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
    // Пропускаем поля, скрытые фильтром типа объекта
    if (typeof isInputVisible === 'function' && !isInputVisible(el)) continue;
    if (el.value.trim() === '') {
      issues.push(f);
    }
  }
  return issues;
}


// ── Обновить Smart Panel: проверка данных ─────────────────────
let _currentIssues = [];
let _issueIndex = 0;

function getBuyerRequiredByDeposit() {
  const depositBYN = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
  const depositUSD = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
  return depositBYN !== '' || depositUSD !== '';
}

function hasRequiredBlockData(blockId) {
  const block = document.getElementById(blockId);
  if (!block) return false;
  return [...block.querySelectorAll('input[type="text"]')].some(el => {
    if (el.readOnly) return false;
    if (typeof isInputVisible === 'function' && !isInputVisible(el)) return false;
    return el.value.trim() !== '';
  });
}

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
  const _hasBuyer = getBuyerRequiredByDeposit();

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
      // Синхронизировано с бейджем заголовка: статус «Заполнено»
      // ставится только когда в блоке нет незаполненных обязательных полей.
      const hasData = hasRequiredBlockData(blockId);
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

// ── Краткая карточка объекта ──────────────────────────────────
function readFieldValue(id) {
  return (document.getElementById(id)?.value || '').trim();
}

function firstFilledValue(...ids) {
  for (const id of ids) {
    const value = readFieldValue(id);
    if (value) return value;
  }
  return '';
}

function normalizeObjectType(value) {
  const raw = (value || '').trim();
  const low = raw.toLowerCase();
  if (low.includes('коммер')) return 'Коммерческая недвижимость';
  if (low.includes('кварт') || low.includes('апартамент') || low.includes('комнат')) return 'Квартира';
  if (low.includes('дом')) return 'Дом';
  return raw || '—';
}

function getObjectTypeIcon(type) {
  const low = (type || '').toLowerCase();
  if (low.includes('коммер')) {
    return `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="11" y="21" width="42" height="33" rx="4" fill="#E8F5EF" stroke="currentColor" stroke-width="3"/><path d="M20 21V13h24v8" stroke="currentColor" stroke-width="3" stroke-linejoin="round"/><path d="M19 31h8M36 31h8M19 40h8M36 40h8M31 54V43h8v11" stroke="#2F7A63" stroke-width="3" stroke-linecap="round"/></svg>`;
  }
  if (low.includes('кварт') || low.includes('апартамент') || low.includes('комнат')) {
    return `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="15" y="9" width="34" height="47" rx="4" fill="#EAF4FF" stroke="currentColor" stroke-width="3"/><path d="M24 19h5M35 19h5M24 29h5M35 29h5M24 39h5M35 39h5M29 56V46h6v10" stroke="#2F7A63" stroke-width="3" stroke-linecap="round"/></svg>`;
  }
  return `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M9 30 32 12l23 18" fill="#DDF4E8"/><path d="M9 30 32 12l23 18" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/><path d="M16 28v26h32V28" fill="#F4FBF8"/><path d="M16 28v26h32V28" stroke="currentColor" stroke-width="3" stroke-linejoin="round"/><path d="M28 54V39h9v15M22 34h9" stroke="#2F7A63" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
}

function formatSummaryArea(value) {
  if (!value) return '';
  return `${value.replace(/\s*(м2|м²|кв\.?\s*м)\s*$/i, '')} м²`;
}

function updateObjectSummary() {
  const card = document.getElementById('object-summary-card');
  if (!card) return;

  const type = normalizeObjectType(readFieldValue('property-Тип объекта'));
  const title = document.getElementById('object-summary-title');
  const address = document.getElementById('object-summary-address');
  const meta = document.getElementById('object-summary-meta');
  const priceByn = document.getElementById('object-summary-price-byn');
  const priceUsd = document.getElementById('object-summary-price-usd');
  const icon = document.getElementById('object-summary-icon');

  if (title) title.textContent = `Объект: ${type}`;
  if (icon) icon.innerHTML = getObjectTypeIcon(type);

  const fullAddress = firstFilledValue('property-Адрес');
  const city = readFieldValue('property-Город');
  const street = readFieldValue('property-Улица');
  const house = readFieldValue('property-Дом');
  const flat = readFieldValue('property-Квартира');
  const addressParts = fullAddress ? [fullAddress] : [city, street, house && `д. ${house}`, flat && `кв. ${flat}`].filter(Boolean);
  if (address) address.textContent = addressParts.join(', ') || 'Адрес не указан';

  const metaItems = [
    formatSummaryArea(readFieldValue('property-Общая площадь')),
    readFieldValue('property-Количество комнат') && `${readFieldValue('property-Количество комнат')} комнаты`,
    readFieldValue('property-Материал стен'),
    readFieldValue('property-Год постройки') && `${readFieldValue('property-Год постройки')} г.п.`,
  ].filter(Boolean);
  if (meta) {
    meta.innerHTML = metaItems.length
      ? metaItems.map((item) => `<span class="object-summary-meta-item">${escHtmlUI(item)}</span>`).join('<span class="object-summary-meta-sep" aria-hidden="true"></span>')
      : '<span class="object-summary-meta-item">Характеристики не указаны</span>';
  }

  const byn = readFieldValue('deal-Стоимость BYN');
  const usd = readFieldValue('deal-Стоимость USD');
  if (priceByn) priceByn.textContent = byn ? `${byn} BYN` : '— BYN';
  if (priceUsd) priceUsd.textContent = usd ? `≈ ${usd} USD` : '≈ — USD';
}

// ── Полное обновление Smart Panel ─────────────────────────────
function refreshUI() {
  updateObjectSummary();
  updateValidationPanel();
  if (typeof updateBlockCompletion === 'function') updateBlockCompletion(null);
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
    saveUIState();
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
const SP_TOGGLES = ['sp-validation-toggle', 'sp-docs-toggle', 'sp-settings-toggle'];

function saveUIState() {
  try {
    const openBlocks = BLOCKS.filter(id => document.getElementById(id)?.classList.contains('ws-block--open'));
    const spExpanded = {};
    SP_TOGGLES.forEach(id => {
      const el = document.getElementById(id);
      if (el) spExpanded[id] = el.getAttribute('aria-expanded') === 'true';
    });
    localStorage.setItem(LS_KEY, JSON.stringify({ openBlocks, spExpanded }));
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
    // Restore sp-panel toggle states
    if (state.spExpanded && typeof state.spExpanded === 'object') {
      SP_TOGGLES.forEach(id => {
        if (!(id in state.spExpanded)) return;
        const toggle = document.getElementById(id);
        const bodyId = id.replace('-toggle', '-body');
        const body   = document.getElementById(bodyId);
        if (!toggle || !body) return;
        const shouldBeExpanded = state.spExpanded[id];
        toggle.setAttribute('aria-expanded', String(shouldBeExpanded));
        body.classList.toggle('sp-block-body--collapsed', !shouldBeExpanded);
        toggle.classList.toggle('sp-block-hdr--collapsed', !shouldBeExpanded);
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
    updateObjectSummary();
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

// ── Умный автофокус по Enter ──────────────────────────────────
// Enter на редактируемом поле → следующее видимое поле.
// Если оно в закрытой секции — секция открывается автоматически.
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter' || e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;

  const input = e.target;
  if (!input.matches('.ws-block input[type="text"]:not([readonly]):not([disabled])')) return;

  // Все видимые редактируемые поля в порядке DOM
  const all = [...document.querySelectorAll(
    '.ws-block input[type="text"]:not([readonly]):not([disabled])'
  )].filter(el => el.offsetParent !== null);

  const idx = all.indexOf(input);
  if (idx === -1 || idx === all.length - 1) return;   // последнее поле — не трогаем

  e.preventDefault();

  const next      = all[idx + 1];
  const nextBlock = next.closest('.ws-block');

  if (nextBlock && !nextBlock.classList.contains('ws-block--open')) {
    // Открываем секцию и ждём завершения анимации
    openBlock(nextBlock.id, true);
    setTimeout(() => next.focus(), 100);
  } else {
    next.focus();
  }
});

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
