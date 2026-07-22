'use strict';
/**
 * notification-center.js — Единый центр уведомлений
 *
 * window.NotificationCenter API:
 *   .show({ message, subtitle, type, duration, actions })
 *   .success(message, opts)
 *   .error(message, opts)
 *   .info(message, opts)
 *   .warn(message, opts)
 *
 * types:    'success' | 'error' | 'info' | 'warn'
 * actions:  [{ label: 'Отменить', onClick: fn }]
 *
 * Не более 3 уведомлений одновременно.
 * Новые уведомления вытесняют старые без визуального скачка.
 */

(function () {

  const MAX_TOASTS   = 3;
  const DEFAULT_DURATION = { success: 3200, error: 5500, info: 3500, warn: 4200 };

  const ICONS = {
    success: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>',
    error:   '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
    info:    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="8.01"/><line x1="12" y1="12" x2="12" y2="16"/></svg>',
    warn:    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
  };

  // Очередь активных уведомлений
  let _queue = [];
  let _container = null;

  function _getContainer() {
    if (!_container) _container = document.getElementById('toast-container');
    return _container;
  }

  function _esc(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function _dismiss(entry) {
    if (entry.dismissed) return;
    entry.dismissed = true;
    clearTimeout(entry.timer);
    // Synchronously remove from queue so the capacity check in show() always terminates.
    // DOM removal is deferred for the exit animation.
    _queue = _queue.filter(q => q !== entry);
    const el = entry.el;
    el.classList.remove('toast--visible');
    el.classList.add('toast--leaving');
    setTimeout(() => el.remove(), 230);
  }

  /**
   * Показать уведомление.
   * @param {Object|string} opts
   * @returns {{ dismiss: Function }}
   */
  function show(opts) {
    if (typeof opts === 'string') opts = { message: opts };
    const {
      message  = '',
      subtitle = null,
      type     = 'success',
      duration = DEFAULT_DURATION[type] ?? 3200,
      actions  = [],
    } = opts;

    const container = _getContainer();
    if (!container) return { dismiss: () => {} };

    // Вытесняем самое старое если исчерпан лимит
    while (_queue.length >= MAX_TOASTS) {
      _dismiss(_queue[0]);
    }

    const el = document.createElement('div');
    el.className = `toast toast-${type} toast--nc`;
    el.setAttribute('role', 'status');
    el.setAttribute('aria-live', 'polite');
    el.setAttribute('aria-atomic', 'true');

    let html = `<span class="toast__icon toast__icon--${type}">${ICONS[type] || ''}</span>`;
    html += `<div class="toast__body">`;
    html += `<span class="toast__msg">${_esc(message)}</span>`;
    if (subtitle) {
      html += `<span class="toast__sub">${_esc(subtitle)}</span>`;
    }
    if (actions.length > 0) {
      html += `<div class="toast__acts">`;
      actions.forEach((a, i) => {
        html += `<button class="toast__act" data-nc-idx="${i}" type="button">${_esc(a.label)}</button>`;
      });
      html += `</div>`;
    }
    html += `</div>`;
    html += `<button class="toast__x" aria-label="Закрыть" type="button" title="Закрыть">`
          + `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`
          + `</button>`;

    el.innerHTML = html;

    const entry = { el, dismissed: false, timer: null };
    _queue.push(entry);

    // Закрыть по кнопке
    el.querySelector('.toast__x').addEventListener('click', () => _dismiss(entry));

    // Привязать action-кнопки
    actions.forEach((a, i) => {
      const btn = el.querySelector(`.toast__act[data-nc-idx="${i}"]`);
      if (btn) {
        btn.addEventListener('click', () => {
          _dismiss(entry);
          try { a.onClick?.(); } catch (_) {}
        });
      }
    });

    container.appendChild(el);

    // Анимация появления (два RAF гарантируют применение начального состояния)
    requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add('toast--visible')));

    // Авто-скрытие
    if (duration > 0) {
      entry.timer = setTimeout(() => _dismiss(entry), duration);
    }

    return { dismiss: () => _dismiss(entry) };
  }

  // ── Публичный API ─────────────────────────────────────────────
  window.NotificationCenter = {
    show,
    success: (msg, opts = {}) => show({ ...opts, message: msg, type: 'success' }),
    error:   (msg, opts = {}) => show({ ...opts, message: msg, type: 'error',  duration: opts.duration ?? DEFAULT_DURATION.error }),
    info:    (msg, opts = {}) => show({ ...opts, message: msg, type: 'info' }),
    warn:    (msg, opts = {}) => show({ ...opts, message: msg, type: 'warn' }),
  };

})();
