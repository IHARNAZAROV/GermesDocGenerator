'use strict';
/**
 * notifications.js — единый центр уведомлений ГермесГарант
 *
 * Публичный API (window.Notify):
 *   Notify.toast(opts | string)   — Toast уведомление
 *   Notify.status(text, mode?)    — Статус-бар ('saving' | 'saved' | '')
 *   Notify.progressStart(opts)    — Начать overlay прогресса
 *   Notify.progressStep(label)    — Добавить шаг (обрабатывается)
 *   Notify.progressDone(done,tot) — Завершить текущий шаг
 *   Notify.progressHide()         — Закрыть overlay
 *   Notify.highlightField(id)     — Плавная подсветка поля + прокрутка
 *
 * toast opts:
 *   type:     'success' | 'warning' | 'error' | 'info'  (default: 'success')
 *   title:    string  — основной текст
 *   body:     string  — дополнительный текст (опционально)
 *   duration: number ms (default: 4200; 0 = не закрывать авто.)
 *   actions:  [{label, onClick, closeOnClick?}]
 *   onUndo:   function — добавляет кнопку «Отменить» на 5 секунд
 */
(function () {

  const MAX_TOASTS    = 3;
  const DEFAULT_DUR   = 4200;

  // ── SVG иконки ──────────────────────────────────────────────
  const ICONS = {
    success: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
    warning: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    error:   '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
    info:    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="8.01"/><line x1="12" y1="12" x2="12" y2="16"/></svg>',
    spinner: '<svg class="nt-spin" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg>',
    check:   '<svg class="nt-check-anim" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
  };

  function _esc(s) {
    return String(s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  // ─────────────────────────────────────────────────────────────
  //  TOAST
  // ─────────────────────────────────────────────────────────────

  function toast(opts) {
    if (typeof opts === 'string') opts = { title: opts };

    const {
      type     = 'success',
      title    = '',
      body     = null,
      duration = DEFAULT_DUR,
      actions  = [],
      onUndo   = null,
    } = opts;

    const container = document.getElementById('toast-container');
    if (!container) return null;

    // Не более MAX_TOASTS одновременно — убираем старейший
    const existing = container.querySelectorAll('.toast');
    if (existing.length >= MAX_TOASTS) _dismiss(existing[0], true);

    // Строим элемент
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.setAttribute('role', 'status');
    el.setAttribute('aria-live', 'polite');

    let html = `<span class="toast-icon" aria-hidden="true">${ICONS[type] || ICONS.info}</span>`;
    html += '<div class="toast-content">';
    if (title) html += `<div class="toast-title">${_esc(title)}</div>`;
    if (body)  html += `<div class="toast-body">${_esc(body)}</div>`;

    const allActions = [...actions];
    if (onUndo) allActions.push({ label: 'Отменить', onClick: onUndo, style: 'undo' });

    if (allActions.length) {
      html += '<div class="toast-actions">';
      allActions.forEach(a => {
        const cls = a.style === 'undo' ? 'toast-action toast-action--undo' : 'toast-action';
        html += `<button class="${cls}">${_esc(a.label)}</button>`;
      });
      html += '</div>';
    }
    html += '</div>';
    html += '<button class="toast-close" aria-label="Закрыть">✕</button>';

    el.innerHTML = html;

    // Привязываем события
    el.querySelectorAll('.toast-action').forEach((btn, i) => {
      const action = allActions[i];
      if (!action) return;
      btn.addEventListener('click', () => {
        action.onClick?.();
        if (action.closeOnClick !== false) _dismiss(el);
      });
    });
    el.querySelector('.toast-close')?.addEventListener('click', () => _dismiss(el));

    // Таймер «Отменить» — 5 секунд
    if (onUndo) {
      const undoBtn = el.querySelector('.toast-action--undo');
      let secs = 5;
      if (undoBtn) {
        const tick = setInterval(() => {
          secs--;
          if (secs <= 0) { clearInterval(tick); _dismiss(el); }
          else undoBtn.textContent = `Отменить (${secs}с)`;
        }, 1000);
        el._undoTimer = tick;
        undoBtn.textContent = `Отменить (${secs}с)`;
      }
    }

    container.appendChild(el);
    requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add('toast-visible')));

    if (duration > 0) {
      el._timer = setTimeout(() => _dismiss(el), duration);
    }

    return el;
  }

  function _dismiss(el, immediate = false) {
    if (!el || el._dismissing) return;
    el._dismissing = true;
    clearTimeout(el._timer);
    clearInterval(el._undoTimer);
    if (immediate) {
      el.remove();
      return;
    }
    el.classList.remove('toast-visible');
    el.classList.add('toast-dismissing');
    setTimeout(() => el.remove(), 220);
  }

  // ─────────────────────────────────────────────────────────────
  //  STATUS BAR
  // ─────────────────────────────────────────────────────────────

  function status(text, mode) {
    const dot  = document.querySelector('.status-dot');
    const span = document.getElementById('status-text');
    if (span) span.textContent = text;
    if (dot) {
      dot.className = 'status-dot';
      if (mode === 'saving') dot.classList.add('status-dot--saving');
      if (mode === 'saved')  dot.classList.add('status-dot--saved');
    }
  }

  // ─────────────────────────────────────────────────────────────
  //  PROGRESS OVERLAY
  // ─────────────────────────────────────────────────────────────

  let _po = null; // active overlay element

  function progressStart(opts) {
    const { title = 'Формирование документов', total = 0 } = opts || {};
    const overlay = document.getElementById('progress-overlay');
    if (!overlay) return;
    _po = overlay;
    _po._total = total;
    _po._done  = 0;
    overlay.querySelector('.po-title').textContent = title;
    overlay.querySelector('.po-steps').innerHTML   = '';
    _setBar(0, total);
    overlay.hidden = false;
    requestAnimationFrame(() => overlay.classList.add('po-visible'));
  }

  function progressStep(label) {
    if (!_po) return;
    const steps = _po.querySelector('.po-steps');
    const item  = document.createElement('div');
    item.className = 'po-step';
    item.innerHTML = `<span class="po-step-icon">${ICONS.spinner}</span><span class="po-step-label">${_esc(label)}</span>`;
    steps.appendChild(item);
    steps.scrollTop = steps.scrollHeight;
  }

  function progressDone(done, total) {
    if (!_po) return;
    // Отмечаем последний шаг как выполненный
    const last = _po.querySelector('.po-steps .po-step:last-child');
    if (last) {
      last.classList.add('po-step--done');
      const icon = last.querySelector('.po-step-icon');
      if (icon) icon.innerHTML = ICONS.check;
    }
    _po._done  = done  ?? _po._total;
    _po._total = total ?? _po._total;
    _setBar(_po._done, _po._total);
  }

  function progressHide() {
    if (!_po) return;
    _setBar(_po._total, _po._total);
    setTimeout(() => {
      _po?.classList.remove('po-visible');
      const el = _po;
      setTimeout(() => { if (el) el.hidden = true; }, 220);
      _po = null;
    }, 300);
  }

  function _setBar(done, total) {
    if (!_po) return;
    const pct   = total > 0 ? Math.min(100, Math.round(done / total * 100)) : 0;
    const fill  = _po.querySelector('.po-fill');
    const count = _po.querySelector('.po-count');
    if (fill)  fill.style.width = pct + '%';
    if (count) count.textContent = total > 0 ? `${done} из ${total}` : '';
  }

  // ─────────────────────────────────────────────────────────────
  //  FIELD HIGHLIGHT
  // ─────────────────────────────────────────────────────────────

  function highlightField(inputId) {
    const el = document.getElementById(inputId);
    if (!el) return;

    // Раскрыть секцию если нужно
    const section = el.closest('.ws-block:not(.ws-block--open)');
    if (section) {
      const hdr = section.querySelector('.ws-block-hdr');
      hdr?.click();
    }

    // Прокрутка с небольшой задержкой (даёт время на открытие аккордеона)
    setTimeout(() => {
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.focus({ preventScroll: true });
      el.classList.add('field-highlight');
      setTimeout(() => el.classList.remove('field-highlight'), 1800);
    }, 120);
  }

  // ─────────────────────────────────────────────────────────────
  //  ПУБЛИЧНЫЙ API
  // ─────────────────────────────────────────────────────────────

  window.Notify = {
    toast,
    status,
    progressStart,
    progressStep,
    progressDone,
    progressHide,
    highlightField,
  };

}());
