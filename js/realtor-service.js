'use strict';

// ============================================================
//  RealtorService — единый источник правды о выбранном риэлтере
//
//  Данные берутся из window.AGENTS_CONFIG (js/agents-config.js),
//  который загружается раньше этого скрипта. Отдельный JSON-файл
//  не нужен — agents-config.js / config/agents.json являются
//  единственным справочником сотрудников.
//
//  Выбор сохраняется в localStorage и восстанавливается при
//  следующем запуске. При первом запуске автоматически выбирается
//  первый сотрудник в списке.
// ============================================================

(function () {
  const STORAGE_KEY = 'gg-selected-realtor-id';

  // ── Данные из уже загруженного AGENTS_CONFIG ──────────────
  const _realtors = (window.AGENTS_CONFIG?.agents || []).filter(a => a.id);

  const _savedId  = (() => {
    try { return localStorage.getItem(STORAGE_KEY); } catch (_) { return null; }
  })();
  const _hasSaved = _realtors.some(r => r.id === _savedId);
  let _currentId  = _hasSaved ? _savedId : (_realtors[0]?.id ?? null);

  let _listeners  = [];

  // ── Helpers ───────────────────────────────────────────────

  function _getInitials(r) {
    const f = (r.firstName  || '').charAt(0).toUpperCase();
    const l = (r.lastName   || '').charAt(0).toUpperCase();
    return (f + l) || '?';
  }

  function _notify(realtor) {
    _listeners.forEach(fn => {
      try { fn(realtor); } catch (_) {}
    });
  }

  // ── Public API ────────────────────────────────────────────

  const RealtorService = {
    /** @returns {Object[]} */
    getAll()     { return _realtors; },

    /** @returns {Object|null} */
    getCurrent() {
      return _realtors.find(r => r.id === _currentId) || _realtors[0] || null;
    },

    /** @param {string} id */
    setCurrentById(id) {
      const found = _realtors.find(r => r.id === id);
      if (!found) return;
      const previous = _currentId;
      _currentId = id;
      try { localStorage.setItem(STORAGE_KEY, id); } catch (_) {}
      _notify(found);
      // Уведомление только при реальной смене риэлтера (не при первичной инициализации)
      if (previous && previous !== id) {
        const fullName = [found.lastName, found.firstName, found.middleName]
          .filter(Boolean).join(' ') || found.name || id;
        window.NotificationCenter?.info('Активный риэлтер изменён', { subtitle: fullName });
      }
    },

    /** @param {function} fn  Вызывается с новым объектом риэлтера при каждом изменении. */
    onChange(fn) { _listeners.push(fn); },

    /** @returns {string} */
    getInitials: _getInitials,
  };

  // ── UI ────────────────────────────────────────────────────

  function _buildAvatar(realtor, sizePx) {
    const wrap = document.createElement('div');
    wrap.className   = 'realtor-avatar';
    wrap.style.width  = sizePx + 'px';
    wrap.style.height = sizePx + 'px';
    wrap.setAttribute('aria-hidden', 'true');

    if (realtor && realtor.photo) {
      const img    = document.createElement('img');
      img.src      = realtor.photo;
      img.alt      = '';
      img.loading  = 'lazy';
      img.decoding = 'async';
      img.onerror  = () => { img.remove(); _appendInitials(wrap, realtor); };
      wrap.appendChild(img);
    } else if (realtor) {
      _appendInitials(wrap, realtor);
    }
    return wrap;
  }

  function _appendInitials(wrap, realtor) {
    const span       = document.createElement('span');
    span.className   = 'realtor-avatar__initials';
    span.textContent = _getInitials(realtor);
    wrap.appendChild(span);
  }

  function _updateTrigger(realtor) {
    const slot      = document.getElementById('realtor-avatar-slot');
    const nameEl    = document.getElementById('realtor-name-trigger');
    const posEl     = document.getElementById('realtor-position-trigger');
    if (!slot) return;

    slot.innerHTML = '';
    if (realtor) slot.appendChild(_buildAvatar(realtor, 34));
    if (nameEl) nameEl.textContent = realtor
      ? `${realtor.firstName} ${realtor.lastName}` : '—';
    if (posEl)  posEl.textContent  = realtor ? (realtor.position || '') : '';
  }

  function _buildMenuItem(realtor, isSelected) {
    const li = document.createElement('li');
    li.className   = 'realtor-item' + (isSelected ? ' realtor-item--selected' : '');
    li.role        = 'option';
    li.tabIndex    = -1;
    li.dataset.id  = realtor.id;
    li.setAttribute('aria-selected', String(isSelected));
    li.setAttribute('aria-label',
      `${realtor.firstName} ${realtor.lastName}${realtor.position ? ', ' + realtor.position : ''}`);

    const avatar = _buildAvatar(realtor, 36);

    const info  = document.createElement('div');
    info.className = 'realtor-item__info';
    const nameEl   = document.createElement('span');
    nameEl.className   = 'realtor-item__name';
    nameEl.textContent = `${realtor.firstName} ${realtor.lastName}`;
    const posEl        = document.createElement('span');
    posEl.className    = 'realtor-item__position';
    posEl.textContent  = realtor.position || '';
    info.appendChild(nameEl);
    info.appendChild(posEl);

    const check  = document.createElement('span');
    check.className = 'realtor-item__check';
    check.setAttribute('aria-hidden', 'true');
    if (isSelected) {
      check.innerHTML =
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor"' +
        ' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">' +
        '<polyline points="20 6 9 17 4 12"/></svg>';
    }

    li.appendChild(avatar);
    li.appendChild(info);
    li.appendChild(check);
    return li;
  }

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

    _realtors.forEach(r => {
      const li = _buildMenuItem(r, r.id === _currentId);
      li.addEventListener('click',        () => { RealtorService.setCurrentById(r.id); _closeMenu(); });
      li.addEventListener('keydown',      e  => _handleItemKey(e));
      li.addEventListener('pointerenter', () => li.focus());
      menu.appendChild(li);
    });
  }

  // ── Open / close ──────────────────────────────────────────

  function _isOpen() {
    const m = document.getElementById('realtor-menu');
    return m && !m.hidden;
  }

  function _openMenu() {
    const dropdown = document.getElementById('realtor-dropdown');
    const menu     = document.getElementById('realtor-menu');
    const trigger  = document.getElementById('realtor-trigger');
    if (!menu) return;

    _renderMenu();
    menu.hidden = false;
    dropdown?.setAttribute('aria-expanded', 'true');
    trigger?.setAttribute('aria-expanded', 'true');

    requestAnimationFrame(() => {
      const sel   = menu.querySelector('.realtor-item--selected');
      const first = menu.querySelector('.realtor-item');
      (sel || first)?.focus();
    });

    setTimeout(() => document.addEventListener('pointerdown', _onOutside), 0);
  }

  function _closeMenu() {
    const dropdown = document.getElementById('realtor-dropdown');
    const menu     = document.getElementById('realtor-menu');
    const trigger  = document.getElementById('realtor-trigger');
    if (!menu) return;

    menu.hidden = false; // briefly visible so animation plays
    menu.classList.add('realtor-menu--closing');
    setTimeout(() => {
      menu.hidden = true;
      menu.classList.remove('realtor-menu--closing');
    }, 160);

    dropdown?.setAttribute('aria-expanded', 'false');
    trigger?.setAttribute('aria-expanded', 'false');
    document.removeEventListener('pointerdown', _onOutside);
    trigger?.focus();
  }

  function _onOutside(e) {
    const dropdown = document.getElementById('realtor-dropdown');
    if (dropdown && !dropdown.contains(e.target)) _closeMenu();
  }

  // ── Keyboard navigation ───────────────────────────────────

  function _handleItemKey(e) {
    const menu  = document.getElementById('realtor-menu');
    if (!menu) return;
    const items = Array.from(menu.querySelectorAll('.realtor-item:not(.realtor-item--empty)'));
    const idx   = items.indexOf(document.activeElement);

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
        document.activeElement?.click();
        break;
      case 'Escape':
      case 'Tab':
        e.preventDefault();
        _closeMenu();
        break;
    }
  }

  // ── Init ──────────────────────────────────────────────────

  function _init() {
    const trigger = document.getElementById('realtor-trigger');
    if (!trigger) return;

    trigger.addEventListener('click', e => {
      e.stopPropagation();
      _isOpen() ? _closeMenu() : _openMenu();
    });
    trigger.addEventListener('keydown', e => {
      if (['ArrowDown', 'Enter', ' '].includes(e.key)) {
        e.preventDefault();
        if (!_isOpen()) _openMenu();
      } else if (e.key === 'Escape') {
        _closeMenu();
      }
    });

    RealtorService.onChange(r => { _updateTrigger(r); _renderMenu(); });
    _updateTrigger(RealtorService.getCurrent());
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _init);
  } else {
    _init();
  }

  window.RealtorService = RealtorService;
}());
