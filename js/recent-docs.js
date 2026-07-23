// ============================================================
//  RecentDocs — история последних документов
//  Подключается после app.js в index.html.
//  Публичный API: window.RecentDocs.push({ name, type, icon, path, label })
// ============================================================
(function () {
  'use strict';

  const LS_KEY          = 'germesRecent_v1';
  const MAX_STORED      = 20;   // максимум записей в localStorage
  const MAX_VISIBLE     = 8;    // строк в дропдауне до скролла
  const SEARCH_THRESHOLD = 5;   // показывать поиск если записей > этого

  // ── SVG иконки ────────────────────────────────────────────
  const ICONS = {
    excel: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 8l4 4-4 4"/><line x1="14" y1="16" x2="18" y2="16"/></svg>`,
    dogovor: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M9 15l2 2 4-4"/></svg>`,
    doverennost: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    zapros: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg>`,
    soglasie: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12l2 2 4-4"/></svg>`,
    rastorzhenie: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="19"/><line x1="15" y1="13" x2="9" y2="19"/></svg>`,
    raspiska: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/></svg>`,
    zadatok: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>`,
    word: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/><line x1="8" y1="9" x2="10" y2="9"/></svg>`,
  };

  // Маппинг ключа шаблона → иконка и цвет
  const ICON_MAP = {
    'excel':              { icon: 'excel',       color: '#155945' },
    'doverennost-pnd':   { icon: 'doverennost', color: '#7C3AED' },
    'zapros-pnd':        { icon: 'zapros',       color: '#0369A1' },
    'zapros-rsc':        { icon: 'zapros',       color: '#0369A1' },
    'soglasie-obrabotka':{ icon: 'soglasie',     color: '#059669' },
    'rastorzhenie':      { icon: 'rastorzhenie', color: '#DC2626' },
    'raspiska-klyuchi':  { icon: 'raspiska',     color: '#B45309' },
    'zadatok-standart':  { icon: 'zadatok',      color: '#D97706' },
    'word':              { icon: 'word',          color: '#2B579A' },
  };

  function resolveIcon(iconKey) {
    const entry = ICON_MAP[iconKey] || { icon: 'dogovor', color: '#2B579A' };
    return { svg: ICONS[entry.icon] || ICONS.word, color: entry.color };
  }

  // ── Storage ────────────────────────────────────────────────
  function load() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return [];
      const data = JSON.parse(raw);
      return Array.isArray(data.entries) ? data.entries : [];
    } catch { return []; }
  }

  function persist(entries) {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ entries }));
    } catch (_) {}
  }

  function push({ name, type, icon, path, label }) {
    let entries = load();

    // Убираем дубликат по пути
    entries = entries.filter(e => e.path !== path);

    const entry = {
      id:         `rd_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      name:       name || path.split(/[/\\]/).pop(),
      type,                       // 'excel' | 'word'
      icon:       icon || type,
      path,
      label:      label || name,
      openedAt:   new Date().toISOString(),
      modifiedAt: new Date().toISOString(),
      pinned:     false,
    };

    // Сортировка: закреплённые в начале, затем остальные
    const pinned   = entries.filter(e => e.pinned);
    const unpinned = entries.filter(e => !e.pinned);

    // Бюджет: не больше MAX_STORED суммарно (закреплённые не вытесняются)
    const budget          = Math.max(0, MAX_STORED - pinned.length - 1);
    const trimmedUnpinned = unpinned.slice(0, budget);

    entries = [...pinned, entry, ...trimmedUnpinned];
    persist(entries);
    renderDropdown();
  }

  function togglePin(id) {
    const entries = load();
    const idx = entries.findIndex(e => e.id === id);
    if (idx < 0) return;
    entries[idx].pinned = !entries[idx].pinned;
    // Пересортировка: закреплённые сверху
    const pinned   = entries.filter(e => e.pinned);
    const unpinned = entries.filter(e => !e.pinned);
    persist([...pinned, ...unpinned]);
    renderDropdown();
  }

  function clearHistory() {
    // Удаляем незакреплённые (закреплённые остаются)
    const pinned = load().filter(e => e.pinned);
    persist(pinned);
    renderDropdown();
  }

  // ── Форматирование времени ─────────────────────────────────
  function plural(n, one, few, many) {
    const m10  = n % 10;
    const m100 = n % 100;
    if (m10 === 1 && m100 !== 11)                        return one;
    if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return few;
    return many;
  }
  function pad2(n) { return String(n).padStart(2, '0'); }

  function relTime(isoStr) {
    const diff = Date.now() - new Date(isoStr).getTime();
    if (diff < 0) return 'Только что';
    const mins = Math.floor(diff / 60000);
    if (mins < 1)  return 'Только что';
    if (mins < 60) return `${mins} ${plural(mins, 'минуту', 'минуты', 'минут')} назад`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24)  return `${hrs} ${plural(hrs, 'час', 'часа', 'часов')} назад`;

    const t              = new Date(isoStr);
    const hhmm           = `${pad2(t.getHours())}:${pad2(t.getMinutes())}`;
    const todayStart     = new Date(); todayStart.setHours(0, 0, 0, 0);
    const yesterdayStart = new Date(todayStart); yesterdayStart.setDate(yesterdayStart.getDate() - 1);
    const itemDay        = new Date(isoStr); itemDay.setHours(0, 0, 0, 0);

    if (itemDay.getTime() === todayStart.getTime())     return `Сегодня ${hhmm}`;
    if (itemDay.getTime() === yesterdayStart.getTime()) return `Вчера ${hhmm}`;
    return `${pad2(t.getDate())}.${pad2(t.getMonth() + 1)}.${t.getFullYear()} ${hhmm}`;
  }

  function fullDateStr(isoStr) {
    const t = new Date(isoStr);
    return `${pad2(t.getDate())}.${pad2(t.getMonth() + 1)}.${t.getFullYear()} ${pad2(t.getHours())}:${pad2(t.getMinutes())}`;
  }

  // ── Утилиты ────────────────────────────────────────────────
  function esc(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function trunc(str, max) {
    return str.length <= max ? str : str.slice(0, max - 1) + '…';
  }

  // ── DOM-элементы ───────────────────────────────────────────
  const btnRecent  = document.getElementById('btn-recent');
  const panel      = document.getElementById('recdocs-panel');
  const listEl     = document.getElementById('recdocs-list');
  const countEl    = document.getElementById('recdocs-count');
  const searchWrap = document.getElementById('recdocs-search-wrap');
  const searchInp  = document.getElementById('recdocs-search');
  const btnClear   = document.getElementById('recdocs-clear');
  const btnShowAll = document.getElementById('recdocs-show-all');
  const btnAllDocs = document.getElementById('recdocs-all-docs');

  // ── Состояние дропдауна ────────────────────────────────────
  let isOpen    = false;
  let searchVal = '';

  function openPanel() {
    isOpen = true;
    searchVal = '';
    if (searchInp) searchInp.value = '';
    panel.hidden = false;
    panel.classList.remove('recdocs-panel--closing');
    if (btnRecent) btnRecent.setAttribute('aria-expanded', 'true');
    renderDropdown();
    // Фокус на поиск если видим
    requestAnimationFrame(() => {
      if (searchInp && !searchWrap.hidden) searchInp.focus();
    });
  }

  function closePanel() {
    if (!isOpen) return;
    isOpen = false;
    if (btnRecent) btnRecent.setAttribute('aria-expanded', 'false');
    panel.classList.add('recdocs-panel--closing');
    panel.addEventListener('animationend', () => {
      panel.classList.remove('recdocs-panel--closing');
      panel.hidden = true;
    }, { once: true });
  }

  function togglePanel() {
    if (isOpen) closePanel(); else openPanel();
  }

  // ── Рендер дропдауна ──────────────────────────────────────
  function renderDropdown() {
    const all   = load();
    const total = all.length;

    if (countEl) countEl.textContent = total > 0 ? `(${total})` : '';
    if (searchWrap) searchWrap.hidden = total <= SEARCH_THRESHOLD;

    if (!listEl) return;

    const query    = searchVal.trim().toLowerCase();
    const filtered = query
      ? all.filter(e =>
          e.name.toLowerCase().includes(query) ||
          (e.label || '').toLowerCase().includes(query) ||
          e.path.toLowerCase().includes(query))
      : all;

    listEl.innerHTML = '';

    if (filtered.length === 0) {
      listEl.appendChild(buildEmpty(
        query ? 'Ничего не найдено' : 'Последние документы пока отсутствуют',
        query ? 'Попробуйте другой запрос' : 'После открытия документов они появятся здесь'
      ));
      return;
    }

    filtered.forEach(entry => {
      listEl.appendChild(buildItem(entry));
    });
  }

  function buildEmpty(title, sub) {
    const el = document.createElement('div');
    el.className = 'recdocs-empty';
    el.innerHTML = `
      <svg class="recdocs-empty-icon" width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
      </svg>
      <span class="recdocs-empty-title">${esc(title)}</span>
      <span class="recdocs-empty-sub">${esc(sub)}</span>`;
    return el;
  }

  function buildItem(entry) {
    const { svg, color } = resolveIcon(entry.icon);
    const time           = relTime(entry.openedAt);
    const pinned         = entry.pinned;

    const item = document.createElement('div');
    item.className = 'recdocs-item' + (pinned ? ' recdocs-item--pinned' : '');
    item.dataset.id = entry.id;

    item.innerHTML = `
      <div class="recdocs-item-accent"></div>
      <div class="recdocs-item-icon" style="color:${color}">${svg}</div>
      <div class="recdocs-item-body">
        <span class="recdocs-item-name">${esc(trunc(entry.name, 40))}</span>
        <span class="recdocs-item-meta">
          ${pinned ? '<span class="recdocs-pinbadge" title="Закреплён">📌</span>' : ''}
          ${esc(time)}
        </span>
      </div>
      <button class="recdocs-pin-btn" title="${pinned ? 'Открепить' : 'Закрепить'}" type="button" tabindex="-1">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="${pinned ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
        </svg>
      </button>`;

    // Клик по карточке — открыть
    item.addEventListener('click', e => {
      if (e.target.closest('.recdocs-pin-btn')) return;
      handleOpen(entry);
    });

    // Клик по звёздочке — закрепить/открепить
    item.querySelector('.recdocs-pin-btn').addEventListener('click', e => {
      e.stopPropagation();
      togglePin(entry.id);
    });

    return item;
  }

  // ── Открыть документ ──────────────────────────────────────
  async function handleOpen(entry) {
    closePanel();

    if (entry.type === 'excel') {
      // Уже открыт тот же файл?
      const cur = document.getElementById('file-path-display')?.value;
      if (cur && cur === entry.path) return;

      try {
        // loadExcelFile определена в app.js (глобальная функция)
        await loadExcelFile(entry.path);
      } catch {
        handleMissing(entry);
      }
    } else {
      // Word-документ — открыть через Electron
      try {
        const opened = await window.electronAPI.openFile(entry.path);
        if (opened && opened.error) handleMissing(entry);
      } catch {
        handleMissing(entry);
      }
    }
  }

  function handleMissing(entry) {
    if (typeof showToast === 'function') {
      showToast(`✖ Документ не найден: ${entry.name}`, 'error');
    }
    const entries = load().filter(e => e.id !== entry.id);
    persist(entries);
    renderDropdown();
  }

  // ── Диалог подтверждения ──────────────────────────────────
  function showConfirm(title, msg, onYes) {
    const overlay = document.createElement('div');
    overlay.className = 'recdocs-confirm-overlay';
    overlay.innerHTML = `
      <div class="recdocs-confirm-box">
        <div class="recdocs-confirm-title">${esc(title)}</div>
        <div class="recdocs-confirm-msg">${esc(msg)}</div>
        <div class="recdocs-confirm-btns">
          <button class="recdocs-btn-cancel" type="button">Отмена</button>
          <button class="recdocs-btn-yes"    type="button">Да</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    const close = () => overlay.remove();
    overlay.querySelector('.recdocs-btn-cancel').addEventListener('click', close);
    overlay.querySelector('.recdocs-btn-yes').addEventListener('click', () => { close(); onYes(); });
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });

    function onEsc(e) { if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onEsc); } }
    document.addEventListener('keydown', onEsc);
    setTimeout(() => overlay.querySelector('.recdocs-btn-cancel')?.focus(), 60);
  }

  // ── Модальное окно «Все документы» ────────────────────────
  function openAllModal() {
    closePanel();

    const overlay = document.createElement('div');
    overlay.className = 'recdocs-modal-overlay';
    overlay.innerHTML = `
      <div class="recdocs-modal" role="dialog" aria-modal="true" aria-label="История документов">
        <div class="recdocs-modal-header">
          <span class="recdocs-modal-title">История документов</span>
          <button class="recdocs-modal-close" type="button" aria-label="Закрыть">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <div class="recdocs-modal-search-wrap">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input class="recdocs-modal-search" id="recdocs-modal-inp" type="text" placeholder="Поиск документа..." />
        </div>
        <div class="recdocs-modal-list" id="recdocs-modal-list"></div>
        <div class="recdocs-modal-footer">
          <button class="recdocs-footer-btn recdocs-footer-btn--danger" id="recdocs-modal-clear" type="button">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
            Очистить историю
          </button>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    const modalList   = overlay.querySelector('#recdocs-modal-list');
    const modalSearch = overlay.querySelector('#recdocs-modal-inp');
    const modalClear  = overlay.querySelector('#recdocs-modal-clear');

    const close = () => overlay.remove();
    overlay.querySelector('.recdocs-modal-close').addEventListener('click', close);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });

    function onEsc(e) { if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onEsc); } }
    document.addEventListener('keydown', onEsc);

    function renderModal(q) {
      const all      = load();
      const filtered = q
        ? all.filter(e =>
            e.name.toLowerCase().includes(q.toLowerCase()) ||
            (e.label || '').toLowerCase().includes(q.toLowerCase()))
        : all;

      modalList.innerHTML = '';
      if (filtered.length === 0) {
        modalList.innerHTML = `<div style="text-align:center;padding:40px 0;color:var(--text-muted);font-size:13px;">${q ? 'Ничего не найдено' : 'История пуста'}</div>`;
        return;
      }
      filtered.forEach(entry => {
        const { svg, color } = resolveIcon(entry.icon);
        const row = document.createElement('div');
        row.className = 'recdocs-modal-item';
        row.innerHTML = `
          <div class="recdocs-item-icon" style="color:${color};flex-shrink:0">${svg}</div>
          <div class="recdocs-modal-item-body">
            <span class="recdocs-item-name">${esc(entry.name)}</span>
            <span class="recdocs-modal-item-path">${esc(entry.path)}</span>
          </div>
          <span class="recdocs-modal-time">${esc(relTime(entry.openedAt))}</span>
          <button class="recdocs-pin-btn" title="${entry.pinned ? 'Открепить' : 'Закрепить'}" type="button">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="${entry.pinned ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
            </svg>
          </button>`;
        row.addEventListener('click', e => {
          if (e.target.closest('.recdocs-pin-btn')) {
            togglePin(entry.id);
            renderModal(modalSearch.value);
            return;
          }
          close();
          handleOpen(entry);
        });
        modalList.appendChild(row);
      });
    }

    modalClear.addEventListener('click', () => {
      showConfirm('Очистить историю?', 'Незакреплённые документы будут удалены из списка.', () => {
        clearHistory();
        renderModal(modalSearch.value);
      });
    });

    modalSearch.addEventListener('input', () => renderModal(modalSearch.value));
    renderModal('');
    setTimeout(() => modalSearch.focus(), 80);
  }

  // ── Привязка событий ──────────────────────────────────────
  if (btnRecent) {
    btnRecent.addEventListener('click', e => {
      e.stopPropagation();
      togglePanel();
    });
  }

  // Закрыть по клику вне панели
  document.addEventListener('click', e => {
    if (!isOpen) return;
    const wrap = document.getElementById('recdocs-wrap');
    if (wrap && !wrap.contains(e.target)) closePanel();
  });

  // Закрыть по Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && isOpen) { e.stopPropagation(); closePanel(); }
  }, true);

  // Поиск в дропдауне
  if (searchInp) {
    searchInp.addEventListener('input', () => {
      searchVal = searchInp.value;
      renderDropdown();
    });
    // Не закрывать панель при клике на поле поиска
    searchInp.addEventListener('click', e => e.stopPropagation());
  }

  // Очистить историю (из дропдауна)
  if (btnClear) {
    btnClear.addEventListener('click', e => {
      e.stopPropagation();
      const all = load();
      if (all.length === 0) return;
      showConfirm(
        'Очистить историю?',
        'Незакреплённые документы будут удалены из списка.',
        clearHistory
      );
    });
  }

  // Показать все (заголовок дропдауна)
  if (btnShowAll) {
    btnShowAll.addEventListener('click', e => {
      e.stopPropagation();
      openAllModal();
    });
  }

  // Все документы (подвал дропдауна)
  if (btnAllDocs) {
    btnAllDocs.addEventListener('click', e => {
      e.stopPropagation();
      openAllModal();
    });
  }

  // ── Публичный API ──────────────────────────────────────────
  window.RecentDocs = { push, renderDropdown };

  // Первый рендер
  renderDropdown();
}());
