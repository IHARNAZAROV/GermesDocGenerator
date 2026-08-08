'use strict';
// ============================================================
//  validation.js — единственный источник правды о правилах
//  обязательных полей.
//
//  Используется Smart Panel (ui-controller.js) и кнопкой
//  «Проверить данные» (app.js).
//
//  API: window.Validator.getValidationIssues({ requireBuyer })
//    requireBuyer = true  — покупатель всегда обязателен
//    requireBuyer = false — покупатель обязателен только при
//                          наличии суммы задатка
//
//  Загружать ПЕРЕД app.js и ui-controller.js.
// ============================================================

(function () {

  // ── Базовые поля сделки ───────────────────────────────────
  const DEAL_FIELDS = [
    { id: 'deal-Стоимость BYN',           label: 'Сделка: Цена BYN',       block: 'ws-deal' },
    { id: 'deal-Номер договора',           label: 'Сделка: Номер договора',  block: 'ws-deal' },
    { id: 'deal-Дата договора',           label: 'Сделка: Дата договора',   block: 'ws-deal' },
    { id: 'deal-Дата окончания договора', label: 'Сделка: Дата окончания',  block: 'ws-deal' },
  ];

  // ── Базовые поля объекта ───────────────────────────────────
  const PROPERTY_BASE_FIELDS = [
    { id: 'property-Тип объекта',           label: 'Объект: Тип объекта',    block: 'ws-property' },
    { id: 'property-Адрес',                 label: 'Объект: Адрес',           block: 'ws-property' },
    { id: 'property-Город',                 label: 'Объект: Город',           block: 'ws-property' },
    { id: 'property-Улица',                 label: 'Объект: Улица',           block: 'ws-property' },
    { id: 'property-Дом',                   label: 'Объект: Дом',             block: 'ws-property' },
    { id: 'property-Этаж',                  label: 'Объект: Этаж',            block: 'ws-property' },
    { id: 'property-Этажность',             label: 'Объект: Этажность',       block: 'ws-property' },
    { id: 'property-Количество комнат',     label: 'Объект: Кол-во комнат',   block: 'ws-property' },
    { id: 'property-Общая площадь',         label: 'Объект: Общая площадь',   block: 'ws-property' },
    { id: 'property-Жилая площадь',         label: 'Объект: Жилая площадь',  block: 'ws-property' },
    { id: 'property-Площадь кухни',         label: 'Объект: Площадь кухни',  block: 'ws-property' },
  ];

  // ── Поля покупателя ────────────────────────────────────────
  const BUYER_FIELDS = [
    { id: 'buyer-Фамилия',                 label: 'Покупатель: Фамилия',       block: 'ws-buyer' },
    { id: 'buyer-Имя',                     label: 'Покупатель: Имя',            block: 'ws-buyer' },
    { id: 'buyer-Отчество',                label: 'Покупатель: Отчество',       block: 'ws-buyer' },
    { id: 'buyer-Дата рождения',           label: 'Покупатель: Дата рождения',  block: 'ws-buyer' },
    { id: 'buyer-Паспорт серия',           label: 'Покупатель: Паспорт серия',  block: 'ws-buyer' },
    { id: 'buyer-Паспорт номер',           label: 'Покупатель: Паспорт номер',  block: 'ws-buyer' },
    { id: 'buyer-Идентификационный номер', label: 'Покупатель: Идент. номер',   block: 'ws-buyer' },
    { id: 'buyer-Кем выдан',              label: 'Покупатель: Кем выдан',      block: 'ws-buyer' },
    { id: 'buyer-Дата выдачи',            label: 'Покупатель: Дата выдачи',    block: 'ws-buyer' },
    { id: 'buyer-Адрес регистрации',       label: 'Покупатель: Адрес регистр.', block: 'ws-buyer' },
  ];

  // ── Полный набор полей собственника N ─────────────────────
  function ownerFields(prefix, n) {
    return [
      { id: `${prefix}-Фамилия`,                 label: `Собств.${n}: Фамилия`,          block: 'ws-owners' },
      { id: `${prefix}-Имя`,                     label: `Собств.${n}: Имя`,               block: 'ws-owners' },
      { id: `${prefix}-Отчество`,                label: `Собств.${n}: Отчество`,          block: 'ws-owners' },
      { id: `${prefix}-Дата рождения`,           label: `Собств.${n}: Дата рождения`,     block: 'ws-owners' },
      { id: `${prefix}-Паспорт серия`,           label: `Собств.${n}: Паспорт серия`,     block: 'ws-owners' },
      { id: `${prefix}-Паспорт номер`,           label: `Собств.${n}: Паспорт номер`,     block: 'ws-owners' },
      { id: `${prefix}-Идентификационный номер`, label: `Собств.${n}: Идент. номер`,      block: 'ws-owners' },
      { id: `${prefix}-Кем выдан`,              label: `Собств.${n}: Кем выдан`,         block: 'ws-owners' },
      { id: `${prefix}-Дата выдачи`,            label: `Собств.${n}: Дата выдачи`,       block: 'ws-owners' },
      { id: `${prefix}-Адрес регистрации`,       label: `Собств.${n}: Адрес регистр.`,   block: 'ws-owners' },
      { id: `${prefix}-Доля собственности`,      label: `Собств.${n}: Доля собственности`, block: 'ws-owners' },
    ];
  }

  // ── Поля представителя собственника ───────────────────────
  function representativeFields(prefix, n) {
    return [
      { id: `${prefix}-Представитель фамилия`,      label: `Собств.${n}: Представитель фамилия`,   block: 'ws-owners' },
      { id: `${prefix}-Представитель имя`,           label: `Собств.${n}: Представитель имя`,        block: 'ws-owners' },
      { id: `${prefix}-Представитель паспорт серия`, label: `Собств.${n}: Представ. паспорт серия`,  block: 'ws-owners' },
      { id: `${prefix}-Представитель паспорт номер`, label: `Собств.${n}: Представ. паспорт номер`,  block: 'ws-owners' },
      { id: `${prefix}-Номер доверенности`,          label: `Собств.${n}: Номер доверен.`,           block: 'ws-owners' },
    ];
  }

  // ── Доп. поля объекта по типу ──────────────────────────────
  function propertyTypeFields(propTypeRaw) {
    const t = (propTypeRaw || '').trim().toLowerCase();
    if (t === 'дом' || t === 'жилой дом') {
      return [
        { id: 'property-Кадастровый номер',   label: 'Объект: Кадастровый №',       block: 'ws-property' },
        { id: 'property-Площадь участка',     label: 'Объект: Площадь участка',     block: 'ws-property' },
        { id: 'property-Форма собственности', label: 'Объект: Форма собственности', block: 'ws-property' },
        { id: 'property-Инвентарный номер',   label: 'Объект: Инвентарный №',       block: 'ws-property' },
      ];
    }
    if (t === 'квартира' || t === 'апартаменты' || t === 'комната') {
      return [
        { id: 'property-Инвентарный номер', label: 'Объект: Инвентарный №', block: 'ws-property' },
      ];
    }
    if (t === 'коммерческая недвижимость') {
      return [
        { id: 'property-Вид коммерческой недвижимости',        label: 'Объект: Вид комм. недвижимости',        block: 'ws-property' },
        { id: 'property-Назначение коммерческой недвижимости', label: 'Объект: Назначение комм. недвижимости', block: 'ws-property' },
      ];
    }
    // Тип неизвестен — проверяем оба идентификатора
    return [
      { id: 'property-Кадастровый номер', label: 'Объект: Кадастровый №', block: 'ws-property' },
      { id: 'property-Инвентарный номер', label: 'Объект: Инвентарный №', block: 'ws-property' },
    ];
  }

  // ── Есть ли хоть какие-то данные собственника ─────────────
  function isOwnerFilledAny(prefix) {
    return ['Фамилия', 'Имя', 'Паспорт серия', 'Паспорт номер'].some(key => {
      const el = document.getElementById(`${prefix}-${key}`);
      return el && el.value.trim() !== '';
    });
  }

  // ── Основная функция ──────────────────────────────────────
  /**
   * Возвращает список незаполненных обязательных полей.
   *
   * @param {object}  [opts]
   * @param {boolean} [opts.requireBuyer=true]
   *   true  — покупатель всегда обязателен (кнопка «Проверить»).
   *   false — покупатель обязателен только при наличии задатка
   *           (Smart Panel — меньше шума в процессе заполнения).
   * @returns {{ id: string, label: string, block: string }[]}
   */
  function getValidationIssues({ requireBuyer = true } = {}) {
    const propTypeRaw = (document.getElementById('property-Тип объекта')?.value || '').trim();

    // Нужен ли покупатель
    let needBuyer = requireBuyer;
    if (!needBuyer) {
      const depositBYN = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
      const depositUSD = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
      needBuyer = depositBYN !== '' || depositUSD !== '';
    }

    const required = [
      ...DEAL_FIELDS,
      ...PROPERTY_BASE_FIELDS,
      ...propertyTypeFields(propTypeRaw),
      ...(needBuyer ? BUYER_FIELDS : []),
    ];

    // Собственники: полный набор для каждого, у кого есть данные
    ['owner1', 'owner2', 'owner3'].forEach((prefix, i) => {
      if (!isOwnerFilledAny(prefix)) return;
      required.push(...ownerFields(prefix, i + 1));

      const hasRepEl = document.getElementById(`${prefix}-Есть представитель`);
      if ((hasRepEl?.value || '').trim().toLowerCase() === 'да') {
        required.push(...representativeFields(prefix, i + 1));
      }
    });

    // Покупатель: поля представителя если выбрано «Да»
    if (needBuyer) {
      const buyerRepEl = document.getElementById('buyer-Есть представитель');
      if ((buyerRepEl?.value || '').trim().toLowerCase() === 'да') {
        required.push(
          { id: 'buyer-Представитель фамилия',      label: 'Покупатель: Представитель фамилия',   block: 'ws-buyer' },
          { id: 'buyer-Представитель имя',           label: 'Покупатель: Представитель имя',        block: 'ws-buyer' },
          { id: 'buyer-Представитель паспорт серия', label: 'Покупатель: Представ. паспорт серия',  block: 'ws-buyer' },
          { id: 'buyer-Представитель паспорт номер', label: 'Покупатель: Представ. паспорт номер',  block: 'ws-buyer' },
          { id: 'buyer-Номер доверенности',          label: 'Покупатель: Номер доверен.',           block: 'ws-buyer' },
        );
      }
    }

    const issues = [];
    const seen   = new Set();

    for (const f of required) {
      if (seen.has(f.id)) { continue; }
      seen.add(f.id);

      const el = document.getElementById(f.id);
      if (!el) continue;
      // Пропускаем поля, скрытые фильтром типа объекта или вкладками
      if (typeof isInputVisible === 'function' && !isInputVisible(el)) continue;
      if (el.value.trim() === '') issues.push(f);
    }

    return issues;
  }

  window.Validator = { getValidationIssues };

}());
