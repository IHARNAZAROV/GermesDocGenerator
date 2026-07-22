'use strict';
/**
 * micro-interactions.js — Микровзаимодействия: документы, индикаторы полей.
 *
 * Чистый UI-слой. Бизнес-логику не изменяет.
 * Требует notification-center.js и app.js (FIELD_MAP) быть загруженными ранее.
 */

(function () {

  // ══════════════════════════════════════════════════════════════
  //  1. BATCH DOCUMENT SELECTION TOASTS
  //     Собирает изменения чекбоксов за 500 мс, затем показывает
  //     одно суммарное уведомление вместо N отдельных.
  // ══════════════════════════════════════════════════════════════

  const _batch = { added: 0, removed: 0, lastLabel: '', timer: null };

  function _getDocLabel(item) {
    // Текст элемента без «W» и символов иконки
    const spans = item.querySelectorAll('span');
    for (const s of spans) {
      const t = s.textContent.trim();
      if (t && t !== 'W') return t;
    }
    return (item.textContent || '').replace(/\bW\b/g, '').trim().slice(0, 50);
  }

  function _flushBatch() {
    const { added, removed, lastLabel } = _batch;
    _batch.added = 0; _batch.removed = 0; _batch.lastLabel = '';

    const NC = window.NotificationCenter;
    if (!NC) return;

    if (added > 0 && removed === 0) {
      if (added === 1 && lastLabel) {
        NC.success(`${lastLabel} добавлен`);
      } else {
        NC.success(`Добавлено документов: ${added}`);
      }
    } else if (removed > 0 && added === 0) {
      if (removed === 1 && lastLabel) {
        NC.info(`${lastLabel} удалён из списка`);
      } else {
        NC.info(`Удалено из списка: ${removed}`);
      }
    } else if (added > 0 || removed > 0) {
      NC.info(`Изменено документов: ${added + removed}`);
    }
  }

  document.addEventListener('change', (e) => {
    const cb = e.target;
    if (!cb.matches('.tpl-item input[type="checkbox"]')) return;
    const item = cb.closest('.tpl-item');
    if (!item || item.classList.contains('tpl-item-disabled')) return;

    _batch.lastLabel = _getDocLabel(item);
    if (cb.checked) _batch.added++;
    else            _batch.removed++;

    clearTimeout(_batch.timer);
    _batch.timer = setTimeout(_flushBatch, 500);
  });

  // ══════════════════════════════════════════════════════════════
  //  2. REQUIRED FIELD COMPLETION INDICATORS  (○ → ✓)
  //     Маленький индикатор рядом с label обязательного поля.
  //     При заполнении поля анимируется ○ → ✓.
  // ══════════════════════════════════════════════════════════════

  // Список обязательных полей (зеркалит REQUIRED_FIELDS из ui-controller.js)
  const REQ_IDS = new Set([
    // Сделка
    'deal-Стоимость BYN', 'deal-Номер договора',
    'deal-Дата договора', 'deal-Дата окончания договора',
    // Объект
    'property-Тип объекта', 'property-Адрес', 'property-Город',
    'property-Улица', 'property-Дом', 'property-Этаж', 'property-Этажность',
    'property-Количество комнат', 'property-Общая площадь',
    'property-Жилая площадь', 'property-Площадь кухни',
    // Продавец
    'seller-Фамилия', 'seller-Имя', 'seller-Отчество',
    'seller-Дата рождения', 'seller-Паспорт серия', 'seller-Паспорт номер',
    'seller-Идентификационный номер', 'seller-Кем выдан',
    'seller-Дата выдачи', 'seller-Адрес регистрации',
    'seller-Является собственником',
    // Покупатель
    'buyer-Фамилия', 'buyer-Имя', 'buyer-Отчество',
    'buyer-Дата рождения', 'buyer-Паспорт серия', 'buyer-Паспорт номер',
    'buyer-Идентификационный номер', 'buyer-Кем выдан',
    'buyer-Дата выдачи', 'buyer-Адрес регистрации',
    // Собственники
    'owner1-Фамилия', 'owner1-Имя', 'owner1-Паспорт серия',
    'owner1-Паспорт номер', 'owner1-Идентификационный номер',
    'owner1-Адрес регистрации',
  ]);

  function _addIndicator(id) {
    const el  = document.getElementById(id);
    if (!el) return;
    const row = document.getElementById('fr-' + id);
    if (!row || row.querySelector('.fi-dot')) return;

    const dot = document.createElement('span');
    dot.className = 'fi-dot';
    dot.setAttribute('aria-hidden', 'true');

    const label = row.querySelector('label');
    if (label) label.appendChild(dot);

    _syncIndicator(id);
  }

  function _syncIndicator(id) {
    const el  = document.getElementById(id);
    const row = document.getElementById('fr-' + id);
    if (!el || !row) return;
    const dot = row.querySelector('.fi-dot');
    if (!dot) return;

    const filled = el.value.trim() !== '';
    if (filled === dot.classList.contains('fi-dot--ok')) return; // no change
    if (filled) {
      dot.classList.add('fi-dot--ok');
    } else {
      dot.classList.remove('fi-dot--ok');
    }
  }

  function _initIndicators() {
    REQ_IDS.forEach(id => _addIndicator(id));
  }

  function _syncAll() {
    REQ_IDS.forEach(id => _syncIndicator(id));
  }

  // Обновлять индикаторы при вводе
  document.getElementById('deal-body')?.addEventListener('input', (e) => {
    if (REQ_IDS.has(e.target.id)) _syncIndicator(e.target.id);
  });

  // Синхронизировать после заполнения/очистки формы
  document.addEventListener('form:populated', () => {
    setTimeout(() => { _initIndicators(); _syncAll(); }, 120);
  });
  document.addEventListener('form:cleared', () => {
    setTimeout(_syncAll, 100);
  });

  // Инициализация при загрузке (FormBuilder уже отработал)
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(_initIndicators, 200));
  } else {
    setTimeout(_initIndicators, 200);
  }

})();
