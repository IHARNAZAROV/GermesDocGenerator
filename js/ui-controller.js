'use strict';
/**
 * ui-controller.js — трёхпанельный интерфейс: навигация, аккордеон, Smart Panel.
 *
 * Не изменяет бизнес-логику. Читает DOM, обновляет индикаторы.
 * Подключается ПОСЛЕ app.js в index.html.
 */

// ── Ключ хранилища состояния ──────────────────────────────────
const LS_KEY = 'germesUI_v2';

// ── Секции аккордеона ─────────────────────────────────────────
const BLOCKS = ['ws-excel', 'ws-deal', 'ws-property', 'ws-seller', 'ws-owners', 'ws-buyer'];

// ── Проверяемые поля: { id, label, block } ────────────────────
// Соответствует логике handleCheckData в app.js
const REQUIRED_FIELDS = [
  // Сделка
  { id: 'deal-Стоимость BYN',            label: 'Цена BYN',               block: 'ws-deal' },
  { id: 'deal-Номер договора',            label: 'Номер договора',          block: 'ws-deal' },
  { id: 'deal-Дата договора',            label: 'Дата договора',           block: 'ws-deal' },
  { id: 'deal-Дата окончания договора',  label: 'Дата окончания',          block: 'ws-deal' },
  // Объект
  { id: 'property-Тип объекта',    label: 'Тип объекта',      block: 'ws-property' },
  { id: 'property-Адрес',          label: 'Адрес',             block: 'ws-property' },
  { id: 'property-Город',          label: 'Город',             block: 'ws-property' },
  { id: 'property-Улица',          label: 'Улица',             block: 'ws-property' },
  { id: 'property-Дом',            label: 'Дом',               block: 'ws-property' },
  { id: 'property-Этаж',           label: 'Этаж',              block: 'ws-property' },
  { id: 'property-Этажность',      label: 'Этажность',         block: 'ws-property' },
  { id: 'property-Количество комнат', label: 'Кол-во комнат', block: 'ws-property' },
  { id: 'property-Общая площадь',  label: 'Общая площадь',    block: 'ws-property' },
  { id: 'property-Жилая площадь',  label: 'Жилая площадь',   block: 'ws-property' },
  { id: 'property-Площадь кухни',  label: 'Площадь кухни',   block: 'ws-property' },
  // Продавец
  { id: 'seller-Фамилия',                 label: 'Продавец: Фамилия',      block: 'ws-seller' },
  { id: 'seller-Имя',                     label: 'Продавец: Имя',           block: 'ws-seller' },
  { id: 'seller-Отчество',                label: 'Продавец: Отчество',      block: 'ws-seller' },
  { id: 'seller-Дата рождения',           label: 'Продавец: Дата рождения', block: 'ws-seller' },
  { id: 'seller-Паспорт серия',           label: 'Продавец: Паспорт серия', block: 'ws-seller' },
  { id: 'seller-Паспорт номер',           label: 'Продавец: Паспорт номер', block: 'ws-seller' },
  { id: 'seller-Идентификационный номер', label: 'Продавец: Идент. номер',  block: 'ws-seller' },
  { id: 'seller-Кем выдан',              label: 'Продавец: Кем выдан',     block: 'ws-seller' },
  { id: 'seller-Дата выдачи',            label: 'Продавец: Дата выдачи',   block: 'ws-seller' },
  { id: 'seller-Адрес регистрации',       label: 'Продавец: Адрес регистр.', block: 'ws-seller' },
  { id: 'seller-Является собственником',  label: 'Продавец: Собственник?', block: 'ws-seller' },
  // Покупатель
  { id: 'buyer-Фамилия',                 label: 'Покупатель: Фамилия',      block: 'ws-buyer' },
  { id: 'buyer-Имя',                     label: 'Покупатель: Имя',           block: 'ws-buyer' },
  { id: 'buyer-Отчество',                label: 'Покупатель: Отчество',      block: 'ws-buyer' },
  { id: 'buyer-Дата рождения',           label: 'Покупатель: Дата рождения', block: 'ws-buyer' },
  { id: 'buyer-Паспорт серия',           label: 'Покупатель: Паспорт серия', block: 'ws-buyer' },
  { id: 'buyer-Паспорт номер',           label: 'Покупатель: Паспорт номер', block: 'ws-buyer' },
  { id: 'buyer-Идентификационный номер', label: 'Покупатель: Идент. номер',  block: 'ws-buyer' },
  { id: 'buyer-Кем выдан',              label: 'Покупатель: Кем выдан',     block: 'ws-buyer' },
  { id: 'buyer-Дата выдачи',            label: 'Покупатель: Дата выдачи',   block: 'ws-buyer' },
  { id: 'buyer-Адрес регистрации',       label: 'Покупатель: Адрес регистр.', block: 'ws-buyer' },
];

// Поля собственника №1 (базовые, всегда проверяются или при наличии)
const OWNER1_FIELDS = [
  { id: 'owner1-Фамилия',                 label: 'Собств.1: Фамилия',       block: 'ws-owners' },
  { id: 'owner1-Имя',                     label: 'Собств.1: Имя',            block: 'ws-owners' },
  { id: 'owner1-Паспорт серия',           label: 'Собств.1: Паспорт серия',  block: 'ws-owners' },
  { id: 'owner1-Паспорт номер',           label: 'Собств.1: Паспорт номер',  block: 'ws-owners' },
  { id: 'owner1-Идентификационный номер', label: 'Собств.1: Идент. номер',   block: 'ws-owners' },
  { id: 'owner1-Адрес регистрации',       label: 'Собств.1: Адрес регистр.', block: 'ws-owners' },
];

// ── Все поля для прогресса (включая необязательные) ───────────
// Считаем все видимые, не-readonly, не-hidden поля
function getAllFormInputs() {
  const body = document.getElementById('deal-body');
  if (!body) return [];
  return [...body.querySelectorAll('input[type="text"]:not([style*="display:none"])')].filter(el => {
    // skip hidden inputs
    if (el.id === 'file-path-display') return false;
    // skip readonly (computed) fields
    if (el.readOnly) return false;
    return true;
  });
}

// ── Вычислить незаполненные обязательные поля ─────────────────
function getIssues() {
  const issues = [];
  const propTypeRaw = (document.getElementById('property-Тип объекта')?.value || '').trim().toLowerCase();
  const isHouse      = propTypeRaw === 'дом' || propTypeRaw === 'жилой дом';
  const isFlat       = propTypeRaw === 'квартира' || propTypeRaw === 'апартаменты' || propTypeRaw === 'комната';
  const isCommercial = propTypeRaw === 'коммерческая недвижимость';

  // Property-type-specific fields
  const extraPropertyFields = [];
  if (isHouse) {
    extraPropertyFields.push(
      { id: 'property-Кадастровый номер',   label: 'Кадастровый №',       block: 'ws-property' },
      { id: 'property-Площадь участка',     label: 'Площадь участка',     block: 'ws-property' },
      { id: 'property-Форма собственности', label: 'Форма собственности', block: 'ws-property' },
    );
  } else if (isFlat) {
    extraPropertyFields.push(
      { id: 'property-Инвентарный номер', label: 'Инвентарный №', block: 'ws-property' },
    );
  } else if (isCommercial) {
    extraPropertyFields.push(
      { id: 'property-Вид коммерческой недвижимости',        label: 'Вид комм. недвижимости',        block: 'ws-property' },
      { id: 'property-Назначение коммерческой недвижимости', label: 'Назначение комм. недвижимости', block: 'ws-property' },
    );
  }

  // Покупатель обязателен только если указана сумма задатка
  const depositBYN = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
  const depositUSD = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
  const hasBuyer   = depositBYN !== '' || depositUSD !== '';

  const baseFields  = hasBuyer ? REQUIRED_FIELDS : REQUIRED_FIELDS.filter(f => f.block !== 'ws-buyer');
  const allRequired = [...baseFields, ...extraPropertyFields];

  // Check owner1 if seller is not owner or owner1 has some data
  const sellerIsOwnerRaw = (document.getElementById('seller-Является собственником')?.value || '').trim().toLowerCase();
  const sellerNotOwner = sellerIsOwnerRaw === 'нет';
  const owner1HasData = OWNER1_FIELDS.some(f => {
    const el = document.getElementById(f.id);
    return el && el.value.trim() !== '';
  });
  if (sellerNotOwner || owner1HasData) {
    allRequired.push(...OWNER1_FIELDS);
  }

  for (const f of allRequired) {
    const el = document.getElementById(f.id);
    if (!el) continue;
    if (el.value.trim() === '') {
      issues.push(f);
    }
  }
  return issues;
}

// calcProgress / updateProgress removed — progress bar feature removed

// ── Обновить Smart Panel: проверка данных ─────────────────────
let _currentIssues = [];
let _issueIndex = 0;

function updateValidationPanel() {
  const issues = getIssues();
  _currentIssues = issues;
  _issueIndex = 0;

  const list    = document.getElementById('sp-validation-list');
  const empty   = document.getElementById('sp-validation-empty');
  const badge   = document.getElementById('sp-issues-count');
  const errNav  = document.getElementById('sp-err-nav');
  if (!list) return;

  // Remove old items (keep empty placeholder)
  [...list.querySelectorAll('.sp-val-item')].forEach(el => el.remove());

  if (issues.length === 0) {
    if (empty) {
      empty.style.display = '';
      empty.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><polyline points="20 6 9 17 4 12"/></svg><span>Все поля заполнены</span>`;
      empty.classList.add('sp-validation-empty--ok');
    }
    if (badge) badge.hidden = true;
    if (errNav) errNav.hidden = true;
  } else {
    if (empty) {
      empty.style.display = 'none';
      empty.classList.remove('sp-validation-empty--ok');
    }
    if (badge) { badge.textContent = issues.length; badge.hidden = false; }
    if (errNav) { errNav.hidden = false; updateErrNavLabel(); }

    // Группируем по блоку
    const byBlock = {};
    issues.forEach(f => {
      if (!byBlock[f.block]) byBlock[f.block] = [];
      byBlock[f.block].push(f);
    });

    const blockLabels = {
      'ws-deal':     'Сделка',
      'ws-property': 'Объект',
      'ws-seller':   'Продавец',
      'ws-owners':   'Собственники',
      'ws-buyer':    'Покупатель',
    };

    Object.entries(byBlock).forEach(([blockId, blockIssues]) => {
      const groupEl = document.createElement('div');
      groupEl.className = 'sp-val-group sp-val-item';

      const groupHdr = document.createElement('div');
      groupHdr.className = 'sp-val-group-hdr';
      groupHdr.textContent = blockLabels[blockId] || blockId;
      groupEl.appendChild(groupHdr);

      blockIssues.forEach(f => {
        const item = document.createElement('button');
        item.type = 'button';
        item.className = 'sp-val-row sp-val-item';
        item.dataset.fieldId = f.id;
        item.dataset.blockId = f.block;
        item.innerHTML = `
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <span>${escHtmlUI(f.label.replace(/^[^:]+:\s*/, ''))}</span>`;
        item.addEventListener('click', () => navigateToField(f.id, f.block));
        groupEl.appendChild(item);
      });

      list.appendChild(groupEl);
    });
  }

  // Обновить бейджи навигации
  updateNavBadges(issues);
}

function updateErrNavLabel() {
  const label = document.getElementById('sp-err-nav-label');
  if (!label) return;
  if (_currentIssues.length === 0) return;
  label.textContent = `${_issueIndex + 1} / ${_currentIssues.length}`;
}

function escHtmlUI(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Переход к полю по клику на ошибку ─────────────────────────
function navigateToField(fieldId, blockId) {
  // 1. Раскрыть блок
  openBlock(blockId);

  // 2. Прокрутить к блоку, затем к полю
  requestAnimationFrame(() => {
    setTimeout(() => {
      const fieldEl = document.getElementById(fieldId);
      if (!fieldEl) return;

      // Для полей собственников — активировать нужную вкладку
      const ownerMatch = fieldId.match(/^(owner[123])-/);
      if (ownerMatch) {
        const tabId = ownerMatch[1];
        activateOwnerTab(tabId);
      }

      // Прокрутка
      const fieldWrap = document.getElementById('fr-' + fieldId) || fieldEl;
      fieldWrap.scrollIntoView({ behavior: 'smooth', block: 'center' });

      // Фокус + мигание подсветки
      setTimeout(() => {
        fieldEl.focus({ preventScroll: true });
        fieldEl.classList.add('field-highlight');
        setTimeout(() => fieldEl.classList.remove('field-highlight'), 1400);
      }, 150);
    }, 260);
  });
}

function activateOwnerTab(tabId) {
  document.querySelectorAll('.tab-btn[data-tab]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });
  document.querySelectorAll('.tab-pane').forEach(pane => {
    pane.classList.toggle('active', pane.id === 'tab-pane-' + tabId);
  });
}

// ── Навигация по ошибкам: Следующая / Предыдущая ─────────────
function goToIssue(delta) {
  if (_currentIssues.length === 0) return;
  _issueIndex = (_issueIndex + delta + _currentIssues.length) % _currentIssues.length;
  updateErrNavLabel();
  const f = _currentIssues[_issueIndex];
  navigateToField(f.id, f.block);

  // Подсветить соответствующую строку в панели
  document.querySelectorAll('.sp-val-row').forEach((row, i) => {
    row.classList.toggle('sp-val-row--active', row.dataset.fieldId === f.id);
  });
}

document.getElementById('sp-err-prev')?.addEventListener('click', () => goToIssue(-1));
document.getElementById('sp-err-next')?.addEventListener('click', () => goToIssue(+1));

// ── Бейджи в навигационной панели ─────────────────────────────
function updateNavBadges(issues) {
  const countByBlock = {};
  issues.forEach(f => {
    countByBlock[f.block] = (countByBlock[f.block] || 0) + 1;
  });

  const blockToNav = {
    'ws-deal':     'deal',
    'ws-property': 'property',
    'ws-seller':   'seller',
    'ws-owners':   'owners',
    'ws-buyer':    'buyer',
  };

  // Нужен ли покупатель — определяем по наличию суммы задатка
  const _depBYN  = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
  const _depUSD  = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
  const _hasBuyer = _depBYN !== '' || _depUSD !== '';

  Object.entries(blockToNav).forEach(([blockId, navKey]) => {
    const badge  = document.getElementById('nav-badge-' + navKey);
    const status = document.getElementById('nav-status-' + navKey);
    const count  = countByBlock[blockId] || 0;
    if (!badge || !status) return;

    // Покупатель: особый режим когда задаток не указан
    if (blockId === 'ws-buyer' && !_hasBuyer) {
      badge.hidden = true;
      badge.className = 'nav-item-badge';
      status.textContent = 'Ещё не найден';
      status.className = 'nav-item-sub';
      return;
    }

    if (count > 0) {
      badge.textContent = count;
      badge.hidden = false;
      badge.className = 'nav-item-badge nav-item-badge--warn';
      status.textContent = `${count} не заполн.`;
      status.className = 'nav-item-sub nav-item-sub--warn';
    } else {
      badge.hidden = true;
      badge.className = 'nav-item-badge';
      // Check if block has any filled fields
      const hasData = hasBlockData(blockId);
      if (hasData) {
        status.textContent = 'Заполнено';
        status.className = 'nav-item-sub nav-item-sub--ok';
      } else {
        status.textContent = '';
        status.className = 'nav-item-sub';
      }
    }
  });

  // Статус документов
  updateDocsNavStatus();
}

function hasBlockData(blockId) {
  const block = document.getElementById(blockId);
  if (!block) return false;
  return [...block.querySelectorAll('input[type="text"]')].some(el => el.value.trim() !== '');
}

function updateDocsNavStatus() {
  const checked = document.querySelectorAll('.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]:checked').length;
  const statusEl = document.getElementById('nav-status-docs');
  const badgeEl  = document.getElementById('nav-badge-docs');
  if (statusEl) {
    statusEl.textContent = checked > 0 ? `${checked} выбрано` : 'Не выбраны';
    statusEl.className = 'nav-item-sub' + (checked > 0 ? ' nav-item-sub--ok' : '');
  }
  if (badgeEl) {
    if (checked > 0) { badgeEl.textContent = checked; badgeEl.hidden = false; badgeEl.className = 'nav-item-badge nav-item-badge--info'; }
    else { badgeEl.hidden = true; }
  }
}

// ── Excel status ───────────────────────────────────────────────
function updateExcelNavStatus(loaded, fileName) {
  const statusEl = document.getElementById('nav-status-excel');
  if (!statusEl) return;
  if (loaded) {
    statusEl.textContent = fileName || 'Загружен';
    statusEl.className = 'nav-item-sub nav-item-sub--ok';
  } else {
    statusEl.textContent = 'Файл не загружен';
    statusEl.className = 'nav-item-sub';
  }
}

// ── Полное обновление Smart Panel ─────────────────────────────
function refreshUI() {
  updateValidationPanel();
  updateDocsNavStatus();
  // Проверяем статус Excel по drop zone
  const dropSuccess = document.getElementById('drop-success');
  const fileNameEl  = document.getElementById('file-name');
  const loaded = dropSuccess && !dropSuccess.classList.contains('dz-hidden');
  updateExcelNavStatus(loaded, fileNameEl?.textContent || '');
}

// ── Аккордеон ─────────────────────────────────────────────────
function openBlock(blockId, scrollIntoView = true) {
  const block = document.getElementById(blockId);
  if (!block) return;
  const outer = document.getElementById(blockId + '-body');
  const hdr   = block.querySelector('.ws-block-hdr');
  if (!outer || !hdr) return;

  block.classList.add('ws-block--open');
  hdr.setAttribute('aria-expanded', 'true');
  outer.classList.add('ws-block-body-outer--open');

  if (scrollIntoView) {
    setTimeout(() => {
      block.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 60);
  }

  // Обновить активный nav item
  setActiveNavItem(blockId);
  saveUIState();
}

function closeBlock(blockId) {
  const block = document.getElementById(blockId);
  if (!block) return;
  const outer = document.getElementById(blockId + '-body');
  const hdr   = block.querySelector('.ws-block-hdr');
  if (!outer || !hdr) return;

  block.classList.remove('ws-block--open');
  hdr.setAttribute('aria-expanded', 'false');
  outer.classList.remove('ws-block-body-outer--open');
}

function toggleBlock(blockId) {
  const block = document.getElementById(blockId);
  if (!block) return;
  if (block.classList.contains('ws-block--open')) {
    closeBlock(blockId);
  } else {
    openBlock(blockId);
  }
  saveUIState();
}

function openBlockExclusive(blockId) {
  // Закрыть все остальные
  BLOCKS.forEach(id => {
    if (id !== blockId) closeBlock(id);
  });
  openBlock(blockId);
}

function setActiveNavItem(blockId) {
  document.querySelectorAll('.nav-item[data-target]').forEach(item => {
    item.classList.toggle('nav-item--active', item.dataset.target === blockId);
  });
}

// ── Обработчики кликов аккордеона ─────────────────────────────
document.querySelectorAll('.ws-block-hdr').forEach(hdr => {
  hdr.addEventListener('click', () => {
    const block = hdr.closest('.ws-block');
    if (!block) return;
    toggleBlock(block.id);
  });

  // Keyboard: Enter / Space
  hdr.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      hdr.click();
    }
  });
});

// ── Обработчики кликов навигации ──────────────────────────────
document.querySelectorAll('.nav-item[data-target]').forEach(item => {
  const target = item.dataset.target;

  function activate() {
    if (target === 'sp-document') {
      // Прокрутить smart panel к блоку документов
      const docBlock = document.getElementById('sp-document');
      docBlock?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('nav-item--active'));
      item.classList.add('nav-item--active');
    } else {
      openBlockExclusive(target);
    }
  }

  item.addEventListener('click', activate);
  item.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); activate(); }
  });
});

// ── Коллапсируемые блоки Smart Panel ──────────────────────────
['sp-validation-toggle', 'sp-docs-toggle', 'sp-settings-toggle'].forEach(id => {
  const toggle = document.getElementById(id);
  if (!toggle) return;
  const bodyId = id.replace('-toggle', '-body');
  const body   = document.getElementById(bodyId);
  if (!body) return;

  toggle.addEventListener('click', () => {
    const isExpanded = toggle.getAttribute('aria-expanded') === 'true';
    toggle.setAttribute('aria-expanded', String(!isExpanded));
    body.classList.toggle('sp-block-body--collapsed', isExpanded);
    toggle.classList.toggle('sp-block-hdr--collapsed', isExpanded);
  });
  toggle.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle.click(); }
  });
});

// ── Поиск шаблонов ────────────────────────────────────────────
const tplSearch = document.getElementById('tpl-search');
if (tplSearch) {
  tplSearch.addEventListener('input', () => {
    const q = tplSearch.value.trim().toLowerCase();
    document.querySelectorAll('.tpl-item').forEach(item => {
      const text = item.textContent.toLowerCase();
      item.style.display = (!q || text.includes(q)) ? '' : 'none';
    });
    // Hide empty group headers
    document.querySelectorAll('.tpl-group-hdr').forEach(hdr => {
      let next = hdr.nextElementSibling;
      let hasVisible = false;
      while (next && !next.classList.contains('tpl-group-hdr') && !next.classList.contains('tpl-divider')) {
        if (next.style.display !== 'none') hasVisible = true;
        next = next.nextElementSibling;
      }
      hdr.style.display = hasVisible ? '' : 'none';
    });
  });
}

// ── Состояние: сохранение / восстановление ───────────────────
function saveUIState() {
  try {
    const openBlocks = BLOCKS.filter(id => document.getElementById(id)?.classList.contains('ws-block--open'));
    localStorage.setItem(LS_KEY, JSON.stringify({ openBlocks }));
  } catch (_) {}
}

function restoreUIState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return;
    const state = JSON.parse(raw);
    if (Array.isArray(state.openBlocks) && state.openBlocks.length > 0) {
      // Close all first
      BLOCKS.forEach(id => closeBlock(id));
      // Re-open saved
      state.openBlocks.forEach(id => {
        if (BLOCKS.includes(id)) openBlock(id, false);
      });
    }
  } catch (_) {}
}

// ── Слушатели событий для обновления UI ──────────────────────

// Обновление при изменении полей формы
document.getElementById('deal-body')?.addEventListener('input', () => {
  // Debounce: обновлять не чаще 1 раза за 120ms
  clearTimeout(UIController._debounce);
  UIController._debounce = setTimeout(() => {
    updateValidationPanel();
  }, 120);
});

// Обновление при изменении выбора шаблонов
document.addEventListener('change', e => {
  if (e.target.closest('.tpl-item')) {
    updateDocsNavStatus();
  }
});

// Обновление при заполнении формы из Excel (custom event от app.js)
document.addEventListener('form:populated', () => {
  setTimeout(refreshUI, 80);
});
document.addEventListener('form:cleared', () => {
  setTimeout(refreshUI, 80);
  updateExcelNavStatus(false, '');
});

// Обновление статуса Excel при изменении drop zone
const dropSuccessEl = document.getElementById('drop-success');
if (dropSuccessEl) {
  const obs = new MutationObserver(() => {
    const loaded = !dropSuccessEl.classList.contains('dz-hidden');
    const fileNameEl = document.getElementById('file-name');
    updateExcelNavStatus(loaded, fileNameEl?.textContent || '');
  });
  obs.observe(dropSuccessEl, { attributes: true, attributeFilter: ['class'] });
}

// ── Публичный API ─────────────────────────────────────────────
window.UIController = {
  refresh: refreshUI,
  openBlock,
  closeBlock,
  openBlockExclusive,
  navigateToField,
  _debounce: null,
};

// ── Инициализация ─────────────────────────────────────────────
(function init() {
  restoreUIState();
  refreshUI();
})();
