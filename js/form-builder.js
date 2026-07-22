'use strict';
/**
 * form-builder.js — динамическая генерация формы из FIELDS_CONFIG
 *
 * Экспортирует window.FormBuilder.buildForm(config):
 *   — вставляет HTML полей в контейнеры секций
 *   — возвращает FIELD_MAP: { "group-Ключ": "group-Ключ" }
 *
 * Добавление нового поля: добавить в Excel → запустить scan-excel.js →
 * js/fields-config.js обновится → перезапустить приложение.
 * Код менять не нужно.
 */

// ── SVG иконка календаря ─────────────────────────────────────
const CAL_SVG = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'
  + '<rect x="3" y="4" width="18" height="18" rx="2"/>'
  + '<line x1="16" y1="2" x2="16" y2="6"/>'
  + '<line x1="8" y1="2" x2="8" y2="6"/>'
  + '<line x1="3" y1="10" x2="21" y2="10"/>'
  + '</svg>';

// ── Контейнеры секций (sectionId → DOM id) ───────────────────
// Изменение layout: редактируйте это отображение и index.html.
const SECTION_CONTAINERS = {
  'property':    'form-section-property',
  'deal-prices': 'form-section-deal-prices',
  'seller':      'form-section-seller',
  'buyer':       'form-section-buyer',
  'owner1':      'tab-pane-owner1',
  'owner2':      'tab-pane-owner2',
  'owner3':      'tab-pane-owner3',
  'extras':      'form-section-extras',
};

// Секции собственников: рендерятся в две колонки (.tab-inner-grid > div)
const OWNER_SECTIONS = new Set(['owner1', 'owner2', 'owner3']);

// Секции с двухколоночным layout через col-свойство поля
const TWO_COL_SECTIONS = new Set(['property', 'extras', 'deal-prices']);

// ── Генераторы HTML ───────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function inputId(groupId, key) {
  return `${groupId}-${key}`;
}

/** Обычный текстовый инпут */
function htmlText(groupId, field) {
  const id = inputId(groupId, field.key);
  const numAttr = field.numeric ? ' data-numeric="true"' : '';
  const reqCls  = field.required ? ' fr-required' : '';
  const reqInd  = field.required ? '<span class="fr-indicator" aria-hidden="true">✓</span>' : '';
  return `<div class="fr fr-sm${reqCls}" id="fr-${id}"><label>${escHtml(field.label)}${reqInd}</label>`
       + `<div class="input-wrap"><input type="text" id="${id}"${numAttr} /></div></div>`;
}

/** Инпут с кнопкой-календарём */
function htmlDate(groupId, field) {
  const id = inputId(groupId, field.key);
  const reqCls = field.required ? ' fr-required' : '';
  const reqInd = field.required ? '<span class="fr-indicator" aria-hidden="true">✓</span>' : '';
  return `<div class="fr fr-sm${reqCls}" id="fr-${id}"><label>${escHtml(field.label)}${reqInd}</label>`
       + `<div class="input-wrap"><input type="text" id="${id}" class="has-cal" />`
       + `<button class="cal-btn" tabindex="-1">${CAL_SVG}</button></div></div>`;
}

/** Кастомный select (выпадающий список) в стиле realtor-dropdown */
function htmlSelect(groupId, field) {
  const id      = inputId(groupId, field.key);
  const opts    = (field.options || []);
  const reqCls  = field.required ? ' fr-required' : '';
  const reqInd  = field.required ? '<span class="fr-indicator" aria-hidden="true">✓</span>' : '';
  const optHtml = opts.map(opt =>
    `<li class="obj-sel-item" role="option" tabindex="-1" data-value="${escHtml(opt)}" aria-selected="false">`
    + `<span class="obj-sel-item__label">${escHtml(opt)}</span>`
    + `<span class="obj-sel-item__check" aria-hidden="true"></span>`
    + `</li>`
  ).join('');

  return `<div class="fr fr-sm${reqCls}" id="fr-${id}"><label>${escHtml(field.label)}${reqInd}</label>`
       + `<div class="input-wrap">`
       + `<div class="obj-type-dropdown" id="osd-${id}">`
       // Скрытый input — хранит значение; getField / setInputValue работают с ним напрямую
       + `<input type="text" id="${id}" style="position:absolute;opacity:0;pointer-events:none;width:0;height:0" tabindex="-1" aria-hidden="true" />`
       + `<button type="button" class="obj-sel-trigger" id="osd-btn-${id}"`
       +   ` aria-haspopup="listbox" aria-expanded="false" aria-controls="osd-menu-${id}">`
       +   `<span class="obj-sel-value" id="osd-val-${id}">— не выбрано —</span>`
       +   `<svg class="obj-sel-chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>`
       + `</button>`
       + `<ul class="obj-sel-menu" id="osd-menu-${id}" role="listbox" hidden>`
       + optHtml
       + `</ul>`
       + `</div>`
       + `</div></div>`;
}

/** Readonly инпут */
function htmlReadonly(groupId, field, title) {
  const id = inputId(groupId, field.key);
  return `<div class="fr fr-sm" id="fr-${id}"><label>${escHtml(field.label)}</label>`
       + `<div class="input-wrap"><input type="text" id="${id}" readonly`
       + (title ? ` title="${escHtml(title)}"` : '') + ` /></div></div>`;
}

/** Инпут BYN с валидацией + computed прописью */
function htmlByn(groupId, bynField, propisField) {
  const bynId = inputId(groupId, bynField.key);
  let html = `<div class="fr fr-sm" id="fr-${bynId}"><label>${escHtml(bynField.label)}</label>`
           + `<div class="byn-field-wrap">`
           + `<div class="input-wrap"><input type="text" id="${bynId}" data-numeric="true" placeholder="например: 105 000.50" /></div>`
           + `<span class="byn-error" id="byn-error" hidden>Введите корректную сумму</span>`
           + `</div></div>`;
  if (propisField) {
    const propisId = inputId(groupId, propisField.key);
    html += `<div class="fr fr-sm" id="fr-${propisId}"><label>${escHtml(propisField.label)}</label>`
          + `<div class="input-wrap"><input type="text" id="${propisId}" readonly`
          + ` title="Вычисляется автоматически из поля «Цена BYN»" /></div></div>`;
  }
  return html;
}

/** Два инпута через «/» (Корпус / Квартира) */
function htmlSlashPair(groupId, mainField, pairField) {
  const mainId = inputId(groupId, mainField.key);
  const pairId = inputId(groupId, pairField.key);
  return `<div class="fr fr-sm" id="fr-${mainId}"><label>${escHtml(mainField.label)}</label>`
       + `<div class="input-wrap" style="gap:4px">`
       + `<input type="text" id="${mainId}" style="width:52px;flex:none" />`
       + `<span style="font-size:11px;color:#888">/</span>`
       + `<input type="text" id="${pairId}" style="width:52px;flex:none" />`
       + `</div></div>`;
}

/** Этаж из N этажей */
function htmlFloorPair(groupId, mainField, pairField) {
  const mainId = inputId(groupId, mainField.key);
  const pairId = inputId(groupId, pairField.key);
  return `<div class="fr fr-sm" id="fr-${mainId}"><label>${escHtml(mainField.label)}</label>`
       + `<div class="input-wrap"><div class="floor-row">`
       + `<input type="text" id="${mainId}" style="width:42px" />`
       + `<span>из</span>`
       + `<input type="text" id="${pairId}" style="width:42px" />`
       + `</div></div></div>`;
}

// ── Двухколоночный рендер (по свойству col поля) ─────────────

function renderFieldsTwoCol(groupId, fields) {
  const byKey = {};
  fields.forEach(f => { byKey[f.key] = f; });

  // Разделяем «первичные» поля по колонкам
  // pairedUnder-поля и computed-propis следуют за родителем — классифицируем позже
  const col1Primary = [], col2Primary = [], fullPrimary = [];

  for (const field of fields) {
    if (field.hidden) continue;
    if (field.pairedUnder) continue;       // следует за родительским
    if (field.type === 'computed-propis') continue; // следует за BYN

    if (field.col === 'full') fullPrimary.push(field);
    else if (field.col === 2)  col2Primary.push(field);
    else                       col1Primary.push(field); // col:1 или без col
  }

  // Строим набор полей для каждой колонки (первичные + зависимые)
  function buildSubset(primaries) {
    const primaryKeys = new Set(primaries.map(f => f.key));
    const subset = [...primaries];
    for (const field of fields) {
      // Дочернее поле пары → в колонку родителя
      if (field.pairedUnder && primaryKeys.has(field.pairedUnder)) {
        subset.push(field);
      }
      // Прописью → в колонку BYN (если BYN в этой колонке)
      if (field.type === 'computed-propis' && primaries.some(f => f.type === 'byn')) {
        subset.push(field);
      }
    }
    return subset;
  }

  const col1Fields = buildSubset(col1Primary);
  const col2Fields = buildSubset(col2Primary);

  let html = '';

  // Двухколоночная сетка
  if (col1Fields.length > 0 || col2Fields.length > 0) {
    html += '<div class="form-two-col">';
    html += `<div class="form-col">${renderFields(groupId, col1Fields)}</div>`;
    html += `<div class="form-col">${renderFields(groupId, col2Fields)}</div>`;
    html += '</div>';
  }

  // Поля на всю ширину (col:"full") — после сетки
  for (const field of fullPrimary) {
    if (field.type === 'computed-propis') {
      html += htmlReadonly(groupId, field, 'Вычисляется автоматически из поля «Цена BYN»');
    } else if (field.type === 'date') {
      html += htmlDate(groupId, field);
    } else if (field.type === 'readonly') {
      html += htmlReadonly(groupId, field);
    } else {
      html += htmlText(groupId, field);
    }
  }

  return html;
}

// ── Рендер массива полей одной группы ────────────────────────

function renderFields(groupId, fields) {
  // Индекс по ключу для быстрого поиска пар
  const byKey = {};
  fields.forEach(f => { byKey[f.key] = f; });

  const rendered = new Set();
  let html = '';

  for (const field of fields) {
    if (rendered.has(field.key)) continue;
    rendered.add(field.key);

    // Скрытые поля — не отображаются в форме, логика работает через них как обычно
    if (field.hidden) continue;

    // Пропускаем "дочерние" поля — они рендерятся внутри "родительского"
    if (field.pairedUnder) continue;

    // Computed-propis рендерится внутри BYN-блока
    if (field.type === 'computed-propis') continue;

    let fieldHtml = '';

    // BYN-поле: рендерим вместе с прописью
    if (field.type === 'byn') {
      const propis = fields.find(f => f.type === 'computed-propis');
      fieldHtml = htmlByn(groupId, field, propis || null);
      if (propis) rendered.add(propis.key);
    }

    // Поле с парой (Корпус/Квартира или Этаж/Этажность)
    else if (field.pairWith) {
      const pair = byKey[field.pairWith];
      if (pair) {
        rendered.add(pair.key);
        if (field.pairStyle === 'floor') {
          fieldHtml = htmlFloorPair(groupId, field, pair);
        } else {
          fieldHtml = htmlSlashPair(groupId, field, pair);
        }
      }
    }

    // Стандартные типы
    else {
      switch (field.type) {
        case 'date':
          fieldHtml = htmlDate(groupId, field);
          break;
        case 'readonly':
          fieldHtml = htmlReadonly(groupId, field);
          break;
        case 'select':
          fieldHtml = htmlSelect(groupId, field);
          break;
        default:
          fieldHtml = htmlText(groupId, field);
      }
    }

    // Если поле привязано к типу объекта — добавляем data-атрибут на обёртку
    if (fieldHtml && field.objectType) {
      fieldHtml = fieldHtml.replace('<div class="fr ', `<div data-object-type="${field.objectType}" class="fr `);
    }

    html += fieldHtml;
  }

  return html;
}

// ── Основная функция: строит форму, возвращает FIELD_MAP ──────

function buildForm(config) {
  if (!config || !Array.isArray(config.groups)) {
    console.error('FormBuilder: некорректный конфиг');
    return {};
  }

  // 1. Группируем поля по секциям
  //    Каждый элемент: { groupId, field }
  const sections = {};

  for (const group of config.groups) {
    const defaultSection = group.defaultSection || group.id;

    for (const field of (group.fields || [])) {
      const section = field.section || defaultSection;
      if (!sections[section]) sections[section] = [];
      sections[section].push({ groupId: group.id, field });
    }
  }

  // 2. Рендерим каждую секцию в соответствующий DOM-контейнер
  for (const [sectionId, entries] of Object.entries(sections)) {
    const containerId = SECTION_CONTAINERS[sectionId];
    if (!containerId) continue; // неизвестная секция — пропускаем

    const container = document.getElementById(containerId);
    if (!container) continue;

    // Для секций собственников: рендерим в две колонки
    if (OWNER_SECTIONS.has(sectionId)) {
      const inner = container.querySelector('.tab-inner-grid');
      if (!inner) continue;

      const cols = inner.querySelectorAll(':scope > div');
      if (cols.length < 2) continue;

      const groupId = entries[0]?.groupId;
      const fields  = entries.map(e => e.field);
      const mid     = Math.ceil(fields.length / 2);

      cols[0].innerHTML = renderFields(groupId, fields.slice(0, mid));
      cols[1].innerHTML = renderFields(groupId, fields.slice(mid));
      continue;
    }

    // Для остальных секций все поля из одной группы (groupId одинаков в секции)
    const groupId = entries[0]?.groupId;
    const fields  = entries.map(e => e.field);

    if (TWO_COL_SECTIONS.has(sectionId)) {
      container.innerHTML = renderFieldsTwoCol(groupId, fields);
    } else {
      container.innerHTML = renderFields(groupId, fields);
    }
  }

  // 3. Строим FIELD_MAP: "groupId-key" → "groupId-key"
  //    Включаем ВСЕ поля (в т.ч. computed и paired) — они нужны для clearAllInputs
  const fieldMap = {};
  for (const group of config.groups) {
    for (const field of (group.fields || [])) {
      const id = inputId(group.id, field.key);
      fieldMap[id] = id;
    }
  }

  // 4. Инициализируем кастомные select-дропдауны
  _initObjTypeDropdowns();

  return fieldMap;
}

// ── Инициализация кастомных select-дропдаунов ─────────────────

function _syncObjSelLabel(wrapper) {
  const inputEl = wrapper.querySelector('input[type="text"]');
  const trigger = wrapper.querySelector('.obj-sel-trigger');
  const valueEl = wrapper.querySelector('.obj-sel-value');
  const menu    = wrapper.querySelector('.obj-sel-menu');
  if (!inputEl || !valueEl) return;
  const val     = (inputEl.value || '').trim();
  const isEmpty = val === '';
  valueEl.textContent = isEmpty ? '— не выбрано —' : val;
  if (trigger) trigger.classList.toggle('obj-sel-trigger--empty', isEmpty);
  if (menu) {
    menu.querySelectorAll('.obj-sel-item').forEach(li => {
      const selected = li.dataset.value === val;
      li.setAttribute('aria-selected', String(selected));
      li.classList.toggle('obj-sel-item--selected', selected);
      const chk = li.querySelector('.obj-sel-item__check');
      if (chk) chk.innerHTML = selected
        ? '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
          + ' stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">'
          + '<polyline points="20 6 9 17 4 12"/></svg>'
        : '';
    });
  }
}

function _initObjTypeDropdowns() {
  document.querySelectorAll('.obj-type-dropdown').forEach(wrapper => {
    const inputEl = wrapper.querySelector('input[type="text"]');
    const trigger = wrapper.querySelector('.obj-sel-trigger');
    const menu    = wrapper.querySelector('.obj-sel-menu');
    if (!trigger || !menu || !inputEl) return;

    // Синхронизируем начальное состояние (placeholder-стиль)
    _syncObjSelLabel(wrapper);

    function isOpen() { return !menu.hidden; }

    function openMenu() {
      menu.hidden = false;
      trigger.setAttribute('aria-expanded', 'true');
      setTimeout(() => document.addEventListener('pointerdown', onOutside), 0);
      requestAnimationFrame(() => {
        const sel   = menu.querySelector('.obj-sel-item--selected');
        const first = menu.querySelector('.obj-sel-item');
        (sel || first)?.focus();
      });
    }

    function closeMenu() {
      menu.classList.add('obj-sel-menu--closing');
      setTimeout(() => {
        menu.hidden = true;
        menu.classList.remove('obj-sel-menu--closing');
      }, 160);
      trigger.setAttribute('aria-expanded', 'false');
      document.removeEventListener('pointerdown', onOutside);
      trigger.focus();
    }

    function selectValue(val) {
      inputEl.value = val;
      _syncObjSelLabel(wrapper);
      // Уведомляем app.js через bubbling input event
      inputEl.dispatchEvent(new Event('input', { bubbles: true }));
      closeMenu();
    }

    function onOutside(e) {
      if (!wrapper.contains(e.target)) closeMenu();
    }

    trigger.addEventListener('click', () => isOpen() ? closeMenu() : openMenu());
    trigger.addEventListener('keydown', e => {
      if (['ArrowDown', 'Enter', ' '].includes(e.key)) {
        e.preventDefault();
        if (!isOpen()) openMenu();
      } else if (e.key === 'Escape') {
        closeMenu();
      }
    });

    menu.querySelectorAll('.obj-sel-item').forEach(li => {
      li.addEventListener('click', () => selectValue(li.dataset.value));
      li.addEventListener('pointerenter', () => li.focus());
      li.addEventListener('keydown', e => {
        const items = Array.from(menu.querySelectorAll('.obj-sel-item'));
        const idx   = items.indexOf(document.activeElement);
        switch (e.key) {
          case 'ArrowDown':
            e.preventDefault();
            items[Math.min(idx + 1, items.length - 1)]?.focus();
            break;
          case 'ArrowUp':
            e.preventDefault();
            if (idx <= 0) { closeMenu(); return; }
            items[Math.max(idx - 1, 0)]?.focus();
            break;
          case 'Enter':
          case ' ':
            e.preventDefault();
            li.click();
            break;
          case 'Escape':
          case 'Tab':
            e.preventDefault();
            closeMenu();
            break;
        }
      });
    });
  });

  // Синхронизация меток при загрузке/очистке формы
  function syncAll() {
    document.querySelectorAll('.obj-type-dropdown').forEach(_syncObjSelLabel);
  }
  document.addEventListener('form:populated', syncAll);
  document.addEventListener('form:cleared',   syncAll);
}

// ── Публичный API ─────────────────────────────────────────────
window.FormBuilder = { buildForm, syncObjTypeDropdowns: () => document.querySelectorAll('.obj-type-dropdown').forEach(_syncObjSelLabel) };
