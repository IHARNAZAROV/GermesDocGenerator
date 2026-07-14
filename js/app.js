'use strict';

// ============================================================
//  DOM references
// ============================================================
const btnChooseFile   = document.getElementById('btn-choose-file');
const filePathDisplay = document.getElementById('file-path-display');
const errorBanner     = document.getElementById('error-banner');
const errorText       = document.getElementById('error-text');
const errorClose      = document.getElementById('error-close');
const loader          = document.getElementById('loader');
const emptyState      = document.getElementById('empty-state');
const cardsGrid       = document.getElementById('cards-grid');

// ============================================================
//  Field-to-input ID map
//  key  = "<block>-<excel field name>"
//  value = the input's id attribute in index.html
// ============================================================
const FIELD_MAP = {
  // СДЕЛКА
  'deal-Номер сделки':              'deal-Номер сделки',
  'deal-Дата договора':             'deal-Дата договора',
  'deal-Тип договора':              'deal-Тип договора',
  'deal-Эксклюзив':                 'deal-Эксклюзив',
  'deal-Количество собственников':  'deal-Количество собственников',
  'deal-Стоимость USD':             'deal-Стоимость USD',
  'deal-Стоимость BYN':             'deal-Стоимость BYN',
  'deal-Стоимость прописью':        'deal-Стоимость прописью',
  'deal-Комиссия агентства':        'deal-Комиссия агентства',
  // Excel field is "Ответственный риэлтер" — mapped to "Агент" display field
  'deal-Ответственный риэлтер':     'deal-agent',

  // ОБЪЕКТ
  'property-Тип объекта':           'property-Тип объекта',
  'property-Адрес':                 'property-Адрес',
  'property-Город':                 'property-Город',
  'property-Район':                 'property-Район',
  'property-Улица':                 'property-Улица',
  'property-Дом':                   'property-Дом',
  'property-Корпус':                'property-Корпус',
  'property-Квартира':              'property-Квартира',
  'property-Этаж':                  'property-Этаж',
  'property-Этажность':             'property-Этажность',
  'property-Количество комнат':     'property-Количество комнат',
  'property-Общая площадь':         'property-Общая площадь',
  'property-Жилая площадь':         'property-Жилая площадь',
  'property-Площадь кухни':         'property-Площадь кухни',
  'property-Кадастровый номер':     'property-Кадастровый номер',
  'property-Инвентарный номер':     'property-Инвентарный номер',

  // ПРОДАВЕЦ
  'seller-Фамилия':                 'seller-Фамилия',
  'seller-Имя':                     'seller-Имя',
  'seller-Отчество':                'seller-Отчество',
  'seller-Дата рождения':           'seller-Дата рождения',
  'seller-Паспорт серия':           'seller-Паспорт серия',
  'seller-Паспорт номер':           'seller-Паспорт номер',
  'seller-Идентификационный номер': 'seller-Идентификационный номер',
  'seller-Кем выдан':               'seller-Кем выдан',
  'seller-Дата выдачи':             'seller-Дата выдачи',
  'seller-Адрес регистрации':       'seller-Адрес регистрации',
  'seller-Телефон':                 'seller-Телефон',
  'seller-Email':                   'seller-Email',
  'seller-Является собственником':  'seller-Является собственником',
  'seller-Номер доверенности':      'seller-Номер доверенности',
  'seller-Дата доверенности':       'seller-Дата доверенности',

  // СОБСТВЕННИК №1
  'owner1-Фамилия':                 'owner1-Фамилия',
  'owner1-Имя':                     'owner1-Имя',
  'owner1-Отчество':                'owner1-Отчество',
  'owner1-Дата рождения':           'owner1-Дата рождения',
  'owner1-Паспорт серия':           'owner1-Паспорт серия',
  'owner1-Паспорт номер':           'owner1-Паспорт номер',
  'owner1-Идентификационный номер': 'owner1-Идентификационный номер',
  'owner1-Кем выдан':               'owner1-Кем выдан',
  'owner1-Дата выдачи':             'owner1-Дата выдачи',
  'owner1-Адрес регистрации':       'owner1-Адрес регистрации',
  'owner1-Телефон':                 'owner1-Телефон',
  'owner1-Доля собственности':      'owner1-Доля собственности',

  // СОБСТВЕННИК №2
  'owner2-Фамилия':                 'owner2-Фамилия',
  'owner2-Имя':                     'owner2-Имя',
  'owner2-Отчество':                'owner2-Отчество',
  'owner2-Дата рождения':           'owner2-Дата рождения',
  'owner2-Паспорт серия':           'owner2-Паспорт серия',
  'owner2-Паспорт номер':           'owner2-Паспорт номер',
  'owner2-Идентификационный номер': 'owner2-Идентификационный номер',
  'owner2-Кем выдан':               'owner2-Кем выдан',
  'owner2-Дата выдачи':             'owner2-Дата выдачи',
  'owner2-Адрес регистрации':       'owner2-Адрес регистрации',
  'owner2-Телефон':                 'owner2-Телефон',
  'owner2-Доля собственности':      'owner2-Доля собственности',

  // СОБСТВЕННИК №3
  'owner3-Фамилия':                 'owner3-Фамилия',
  'owner3-Имя':                     'owner3-Имя',
  'owner3-Отчество':                'owner3-Отчество',
  'owner3-Дата рождения':           'owner3-Дата рождения',
  'owner3-Паспорт серия':           'owner3-Паспорт серия',
  'owner3-Паспорт номер':           'owner3-Паспорт номер',
  'owner3-Идентификационный номер': 'owner3-Идентификационный номер',
  'owner3-Кем выдан':               'owner3-Кем выдан',
  'owner3-Дата выдачи':             'owner3-Дата выдачи',
  'owner3-Адрес регистрации':       'owner3-Адрес регистрации',
  'owner3-Телефон':                 'owner3-Телефон',
  'owner3-Доля собственности':      'owner3-Доля собственности',

  // ПОКУПАТЕЛЬ
  'buyer-Фамилия':                  'buyer-Фамилия',
  'buyer-Имя':                      'buyer-Имя',
  'buyer-Отчество':                 'buyer-Отчество',
  'buyer-Дата рождения':            'buyer-Дата рождения',
  'buyer-Паспорт серия':            'buyer-Паспорт серия',
  'buyer-Паспорт номер':            'buyer-Паспорт номер',
  'buyer-Идентификационный номер':  'buyer-Идентификационный номер',
  'buyer-Кем выдан':                'buyer-Кем выдан',
  'buyer-Дата выдачи':              'buyer-Дата выдачи',
  'buyer-Адрес регистрации':        'buyer-Адрес регистрации',
  'buyer-Телефон':                  'buyer-Телефон',
  'buyer-Email':                    'buyer-Email',
  'buyer-Семейное положение':       'buyer-Семейное положение',
};

// ============================================================
//  Helpers
// ============================================================
function showError(message) {
  errorText.textContent = message;
  errorBanner.hidden = false;
}

function hideError() {
  errorBanner.hidden = true;
  errorText.textContent = '';
}

function showLoader() { loader.hidden = false; }
function hideLoader() { loader.hidden = true; }

function setInputValue(inputId, value) {
  const el = document.getElementById(inputId);
  if (!el) return;
  const v = (value !== null && value !== undefined) ? String(value).trim() : '';
  el.value = v;
  el.setAttribute('data-empty', v === '' ? 'true' : 'false');
  el.placeholder = v === '' ? '—' : '';
}

function clearAllInputs() {
  Object.values(FIELD_MAP).forEach((id) => setInputValue(id, ''));
}

// ============================================================
//  Populate form from parsed data object
// ============================================================
function populateForm(data) {
  clearAllInputs();

  const blockKeys = ['deal', 'property', 'seller', 'owner1', 'owner2', 'owner3', 'buyer'];

  blockKeys.forEach((block) => {
    const blockData = data[block];
    if (!blockData) return;
    Object.entries(blockData).forEach(([fieldName, value]) => {
      const mapKey = `${block}-${fieldName}`;
      const inputId = FIELD_MAP[mapKey];
      if (inputId) {
        setInputValue(inputId, value);
      }
    });
  });
}

// ============================================================
//  Main flow
// ============================================================
async function handleChooseFile() {
  hideError();

  let filePath;
  try {
    filePath = await window.electronAPI.openFileDialog();
  } catch (err) {
    showError('Не удалось открыть диалог выбора файла: ' + err.message);
    return;
  }

  if (!filePath) return; // user cancelled

  filePathDisplay.textContent = filePath;
  showLoader();

  let data;
  try {
    data = await window.electronAPI.readExcel(filePath);
  } catch (err) {
    hideLoader();
    showError('Ошибка при чтении файла: ' + err.message);
    return;
  }

  hideLoader();

  if (!data || typeof data !== 'object') {
    showError('Файл прочитан, но данные не получены. Проверьте формат файла.');
    return;
  }

  populateForm(data);

  emptyState.hidden = true;
  cardsGrid.hidden = false;
}

// ============================================================
//  Event listeners
// ============================================================
btnChooseFile.addEventListener('click', handleChooseFile);
errorClose.addEventListener('click', hideError);
