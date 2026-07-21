'use strict';

// ============================================================
//  RealtorService — единый источник правды о выбранном риэлтере
//  Загружает data/realtors.json, хранит выбор в localStorage,
//  уведомляет подписчиков об изменении.
// ============================================================

(function () {
  const STORAGE_KEY = 'gg-selected-realtor-id';

  let _realtors   = [];
  let _currentId  = null;
  let _listeners  = [];
  let _readyPromise = null;

  // ── Helpers ───────────────────────────────────────────────

  function _getInitials(r) {
    const f = (r.firstName  || '').charAt(0).toUpperCase();
    const l = (r.lastName   || '').charAt(0).toUpperCase();
    return (f + l) || '?';
  }

  function _notify(realtor) {
    _listeners.forEach(fn => {
      try { fn(realtor); } catch (e) { /* listener must not crash service */ }
    });
  }

  // ── Public API ────────────────────────────────────────────

  const RealtorService = {
    /** @returns {Object[]} */
    getAll()    { return _realtors; },

    /** @returns {Object|null} */
    getCurrent() {
      return _realtors.find(r => r.id === _currentId) || _realtors[0] || null;
    },

    /** @param {string} id */
    setCurrentById(id) {
      const found = _realtors.find(r => r.id === id);
      if (!found) return;
      _currentId = id;
      try { localStorage.setItem(STORAGE_KEY, id); } catch (_) {}
      _notify(found);
    },

    /** @param {function} fn  Called with the new realtor object on every change. */
    onChange(fn) { _listeners.push(fn); },

    /** @returns {Promise<void>}  Resolves once JSON is loaded and UI is ready. */
    ready() { return _readyPromise; },

    /** @returns {string}  Initials for a realtor record (e.g. "ОТ"). */
    getInitials: _getInitials,
  };

  // ── Data loading ──────────────────────────────────────────

  _readyPromise = fetch('./data/realtors.json')
    .then(res => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.json();
    })
    .then(data => {
      _realtors = Array.isArray(data) ? data : [];

      const savedId = (() => {
        try { return localStorage.getItem(STORAGE_KEY); } catch (_) { return null; }
      })();
      const found = _realtors.find(r => r.id === savedId);
      _currentId  = found ? savedId : (_realtors[0]?.id ?? null);

      _initUI();
    })
    .catch(err => {
      console.warn('[RealtorService] Не удалось загрузить data/realtors.json:', err);
      _initUI(); // render empty state gracefully
    });

  // ── UI ────────────────────────────────────────────────────

  /** Build an avatar element (img or initials fallback). */
  function _buildAvatar(realtor, size) {
    const wrap = document.createElement('div');
    wrap.className  = 'realtor-avatar';
    wrap.style.width  = size + 'px';
    wrap.style.height = size + 'px';
    wrap.setAttribute('aria-hidden', 'true');

    if (realtor && realtor.photo) {
      const img = document.createElement('img');
      img.src     = realtor.photo;
      img.alt     = '';
      img.loading = 'lazy';
      img.decoding = 'async';
      img.onerror = () => {
        // replace with initials on load failure
        img.remove();
        _appendInitials(wrap, realtor);
      };
      wrap.appendChild(img);
    } else if (realtor) {
      _appendInitials(wrap, realtor);
    }
    return wrap;
  }

  function _appendInitials(wrap, realtor) {
    const span = document.createElement('span');
    span.className   = 'realtor-avatar__initials';
    span.textContent = _getInitials(realtor);
    wrap.appendChild(span);
  }

  /** Update the trigger button to reflect the currently selected realtor. */
  function _updateTrigger(realtor) {
    const trigger      = document.getElementById('realtor-trigger');
    const avatarSlot   = document.getElementById('realtor-avatar-slot');
    const nameEl       = document.getElementById('realtor-name-trigger');
    const positionEl   = document.getElementById('realtor-position-trigger');
    if (!trigger) return;

    if (avatarSlot) {
      avatarSlot.innerHTML = '';
      if (realtor) avatarSlot.appendChild(_buildAvatar(realtor, 36));
    }
    if (nameEl)     nameEl.textContent     = realtor ? `${realtor.firstName} ${realtor.lastName}` : '—';
    if (positionEl) positionEl.textContent = realtor ? (realtor.position || '') : '';
  }

  /** Build and return a menu item element. */
  function _buildMenuItem(realtor, isSelected, focusIndex) {
    const li = document.createElement('li');
    li.className    = 'realtor-item' + (isSelected ? ' realtor-item--selected' : '');
    li.role         = 'option';
    li.tabIndex     = focusIndex === 0 ? 0 : -1;
    li.dataset.id   = realtor.id;
    li.setAttribute('aria-selected', String(isSelected));
    li.setAttribute('aria-label', `${realtor.firstName} ${realtor.lastName}, ${realtor.position || ''}`);

    const avatar = _buildAvatar(realtor, 36);
    avatar.className += ' realtor-item__avatar';

    const info  = document.createElement('div');
    info.className  = 'realtor-item__info';
    const nameEl    = document.createElement('span');
    nameEl.className = 'realtor-item__name';
    nameEl.textContent = `${realtor.firstName} ${realtor.lastName}`;
    const posEl     = document.createElement('span');
    posEl.className = 'realtor-item__position';
    posEl.textContent = realtor.position || '';
    info.appendChild(nameEl);
    info.appendChild(posEl);

    const check = document.createElement('span');
    check.className = 'realtor-item__check';
    check.setAttribute('aria-hidden', 'true');
    check.innerHTML = isSelected
      ? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'
      : '';

    li.appendChild(avatar);
    li.appendChild(info);
    li.appendChild(check);
    return li;
  }

  /** Render (or re-render) the dropdown menu items. */
  function _renderMenu() {
    const menu = document.getElementById('realtor-menu');
    if (!menu) return;
    menu.innerHTML = '';

    if (_realtors.length === 0) {
      const empty = document.createElement('li');
      empty.className   = 'realtor-item realtor-item--empty';
      empty.textContent = 'Список риэлтеров пуст';
      menu.appendChild(empty);
      return;
    }

    _realtors.forEach((r, i) => {
      const isSelected = r.id === _currentId;
      const li = _buildMenuItem(r, isSelected, i);

      li.addEventListener('click', () => {
        RealtorService.setCurrentById(r.id);
        _closeMenu();
      });
      li.addEventListener('keydown', e => _handleItemKey(e));
      li.addEventListener('pointerenter', () => li.focus());

      menu.appendChild(li);
    });
  }

  // ── Menu open / close ─────────────────────────────────────

  function _openMenu() {
    const dropdown = document.getElementById('realtor-dropdown');
    const menu     = document.getElementById('realtor-menu');
    const trigger  = document.getElementById('realtor-trigger');
    if (!dropdown || !menu) return;

    _renderMenu();
    menu.hidden = false;
    dropdown.setAttribute('aria-expanded', 'true');
    trigger?.setAttribute('aria-expanded', 'true');

    // Focus selected item (or first)
    requestAnimationFrame(() => {
      const selected = menu.querySelector('.realtor-item--selected');
      const first    = menu.querySelector('.realtor-item');
      (selected || first)?.focus();
    });

    // Close on outside click
    setTimeout(() => document.addEventListener('click', _onOutsideClick), 0);
    setTimeout(() => document.addEventListener('pointerdown', _onOutsideClick), 0);
  }

  function _closeMenu() {
    const dropdown = document.getElementById('realtor-dropdown');
    const menu     = document.getElementById('realtor-menu');
    const trigger  = document.getElementById('realtor-trigger');
    if (!menu) return;

    menu.hidden = true;
    dropdown?.setAttribute('aria-expanded', 'false');
    trigger?.setAttribute('aria-expanded', 'false');

    document.removeEventListener('click',       _onOutsideClick);
    document.removeEventListener('pointerdown', _onOutsideClick);

    trigger?.focus();
  }

  function _isMenuOpen() {
    const menu = document.getElementById('realtor-menu');
    return menu && !menu.hidden;
  }

  function _onOutsideClick(e) {
    const dropdown = document.getElementById('realtor-dropdown');
    if (dropdown && !dropdown.contains(e.target)) _closeMenu();
  }

  // ── Keyboard navigation ───────────────────────────────────

  function _handleItemKey(e) {
    const menu  = document.getElementById('realtor-menu');
    if (!menu) return;
    const items = Array.from(menu.querySelectorAll('.realtor-item:not(.realtor-item--empty)'));
    const idx   = items.indexOf(e.currentTarget);

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        items[Math.min(idx + 1, items.length - 1)]?.focus();
        break;
      case 'ArrowUp':
        e.preventDefault();
        if (idx <= 0) { _closeMenu(); break; }
        items[Math.max(idx - 1, 0)]?.focus();
        break;
      case 'Enter':
      case ' ':
        e.preventDefault();
        e.currentTarget.click();
        break;
      case 'Escape':
      case 'Tab':
        e.preventDefault();
        _closeMenu();
        break;
    }
  }

  // ── Initialize DOM ────────────────────────────────────────

  function _initUI() {
    // Trigger button behaviour
    const trigger = document.getElementById('realtor-trigger');
    if (trigger) {
      trigger.addEventListener('click', e => {
        e.stopPropagation();
        _isMenuOpen() ? _closeMenu() : _openMenu();
      });
      trigger.addEventListener('keydown', e => {
        if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          if (!_isMenuOpen()) _openMenu();
        } else if (e.key === 'Escape') {
          _closeMenu();
        }
      });
    }

    // Subscribe to changes → update trigger + re-render menu
    RealtorService.onChange(realtor => {
      _updateTrigger(realtor);
      _renderMenu(); // refresh checkmarks
    });

    // Initial render
    _updateTrigger(RealtorService.getCurrent());
  }

  // ── Export ────────────────────────────────────────────────
  window.RealtorService = RealtorService;
}());
