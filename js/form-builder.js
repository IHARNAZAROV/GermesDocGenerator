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
  return `<div class="fr fr-sm" id="fr-${id}"><label>${escHtml(field.label)}</label>`
       + `<div class="input-wrap"><input type="text" id="${id}" /></div></div>`;
}

/** Инпут с кнопкой-календарём */
function htmlDate(groupId, field) {
  const id = inputId(groupId, field.key);
  return `<div class="fr fr-sm" id="fr-${id}"><label>${escHtml(field.label)}</label>`
       + `<div class="input-wrap"><input type="text" id="${id}" class="has-cal" />`
       + `<button class="cal-btn" tabindex="-1">${CAL_SVG}</button></div></div>`;
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
           + `<div class="input-wrap"><input type="text" id="${bynId}" placeholder="например: 105000.50" /></div>`
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

    // Пропускаем "дочерние" поля — они рендерятся внутри "родительского"
    if (field.pairedUnder) continue;

    // Computed-propis рендерится внутри BYN-блока
    if (field.type === 'computed-propis') continue;

    // BYN-поле: рендерим вместе с прописью
    if (field.type === 'byn') {
      const propis = fields.find(f => f.type === 'computed-propis');
      html += htmlByn(groupId, field, propis || null);
      if (propis) rendered.add(propis.key);
      continue;
    }

    // Поле с парой (Корпус/Квартира или Этаж/Этажность)
    if (field.pairWith) {
      const pair = byKey[field.pairWith];
      if (pair) {
        rendered.add(pair.key);
        if (field.pairStyle === 'floor') {
          html += htmlFloorPair(groupId, field, pair);
        } else {
          html += htmlSlashPair(groupId, field, pair);
        }
        continue;
      }
    }

    // Стандартные типы
    switch (field.type) {
      case 'date':
        html += htmlDate(groupId, field);
        break;
      case 'readonly':
        html += htmlReadonly(groupId, field);
        break;
      default:
        html += htmlText(groupId, field);
    }
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
    container.innerHTML = renderFields(groupId, fields);
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

  return fieldMap;
}

// ── Публичный API ─────────────────────────────────────────────
window.FormBuilder = { buildForm };
