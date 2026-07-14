'use strict';

// ============================================================
//  DOM references
// ============================================================
const btnChooseFile   = document.getElementById('btn-choose-file');
const btnSave         = document.getElementById('btn-save');
const btnSaveAs       = document.getElementById('btn-save-as');
const filePathDisplay = document.getElementById('file-path-display');
const errorBanner     = document.getElementById('error-banner');
const errorText       = document.getElementById('error-text');
const errorClose      = document.getElementById('error-close');
const loader          = document.getElementById('loader');
const emptyState      = document.getElementById('empty-state');
const cardsGrid       = document.getElementById('cards-grid');
const toastContainer  = document.getElementById('toast-container');

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

// Reverse map: inputId → mapKey (for looking up rowMap entries when saving)
const SAVE_MAP = Object.fromEntries(
  Object.entries(FIELD_MAP).map(([mapKey, inputId]) => [inputId, mapKey])
);

// ============================================================
//  Cards auto-expanded after Excel load
// ============================================================
const AUTO_OPEN_CARDS = new Set(['deal', 'property', 'buyer']);

// In-session collapse state
const cardOpenState = {};

// ============================================================
//  Application state
// ============================================================
let currentFilePath = null; // path of the currently open file
let rowMap          = {};   // mapKey → Excel row number  (from _rowMap)
let originalValues  = {};   // inputId → value at load / last save
let dirtyInputIds   = new Set(); // set of inputIds that have been changed

// ============================================================
//  Universal helpers
// ============================================================

/** Returns true when a value should be treated as empty. */
function isFieldEmpty(value) {
  if (value === null || value === undefined) return true;
  return String(value).trim() === '';
}

/**
 * Returns true when every input inside a card section is empty.
 * @param {string} cardId – e.g. 'deal', 'owner2'
 */
function isSectionEmpty(cardId) {
  const card = document.getElementById('card-' + cardId);
  if (!card) return true;
  for (const input of card.querySelectorAll('input[type="text"]')) {
    if (!isFieldEmpty(input.value)) return false;
  }
  return true;
}

// ============================================================
//  Toast notifications
// ============================================================

/**
 * Show a self-dismissing in-app notification.
 * @param {string} message
 * @param {'success'|'error'} type
 */
function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  toastContainer.appendChild(toast);

  requestAnimationFrame(() => {
    requestAnimationFrame(() => toast.classList.add('toast-visible'));
  });

  setTimeout(() => {
    toast.classList.remove('toast-visible');
    setTimeout(() => toast.remove(), 320);
  }, 3200);
}

// ============================================================
//  Error banner helpers
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

// ============================================================
//  Input population
// ============================================================
function setInputValue(inputId, value) {
  const el = document.getElementById(inputId);
  if (!el) return;
  el.value = isFieldEmpty(value) ? '' : String(value).trim();
}

function clearAllInputs() {
  Object.values(FIELD_MAP).forEach((id) => setInputValue(id, ''));
}

// ============================================================
//  Dirty-state tracking
// ============================================================

function updateDirtyState() {
  const hasDirty = dirtyInputIds.size > 0;
  btnSave.disabled = !hasDirty;
  window.electronAPI.notifyDirtyChange(hasDirty);
}

/**
 * Called on every input `input` event.
 * Compares current value against the saved baseline.
 */
function onInputChange(inputId, currentValue) {
  const original = originalValues[inputId] ?? '';
  const current  = currentValue.trim();

  const input = document.getElementById(inputId);
  if (!input) return;

  if (current !== original) {
    dirtyInputIds.add(inputId);
    input.classList.add('input-dirty');
  } else {
    dirtyInputIds.delete(inputId);
    input.classList.remove('input-dirty');
  }

  updateDirtyState();
}

/** Snapshot current values as the new baseline and clear all dirty markers. */
function commitCurrentValues() {
  for (const inputId of Object.values(FIELD_MAP)) {
    const el = document.getElementById(inputId);
    if (!el) continue;
    originalValues[inputId] = el.value.trim();
    el.classList.remove('input-dirty');
  }
  dirtyInputIds.clear();
  updateDirtyState();
}

// Attach change listeners to every input once at startup
for (const inputId of Object.values(FIELD_MAP)) {
  const el = document.getElementById(inputId);
  if (!el) continue;
  el.addEventListener('input', () => onInputChange(inputId, el.value));
}

// ============================================================
//  Build updates map for writing to Excel
//  Returns: { [rowNumber]: value }
// ============================================================
function buildUpdates() {
  const updates = {};
  for (const [inputId, mapKey] of Object.entries(SAVE_MAP)) {
    const rowNum = rowMap[mapKey];
    if (rowNum === undefined) continue;
    const el = document.getElementById(inputId);
    if (el) updates[rowNum] = el.value.trim();
  }
  return updates;
}

// ============================================================
//  Build default "Save As" filename
//  Format: Сделка_<ФамилияПродавца>_<ДатаДоговора>.xlsx
// ============================================================
function buildDefaultSaveAsName() {
  const lastName = (document.getElementById('seller-Фамилия')?.value || '').trim();
  const dealDate = (document.getElementById('deal-Дата договора')?.value || '').trim();

  // Convert DD.MM.YYYY → YYYY-MM-DD for a clean filename
  let dateStr = dealDate;
  const dateParts = dealDate.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (dateParts) dateStr = `${dateParts[3]}-${dateParts[2]}-${dateParts[1]}`;

  const parts = ['Сделка'];
  if (lastName) parts.push(lastName);
  if (dateStr)  parts.push(dateStr);
  return parts.join('_') + '.xlsx';
}

// ============================================================
//  Save — overwrite the currently open file
// ============================================================
async function handleSave() {
  if (!currentFilePath || dirtyInputIds.size === 0) return;

  const updates = buildUpdates();
  try {
    await window.electronAPI.writeExcel(currentFilePath, currentFilePath, updates);
    commitCurrentValues();
    showToast('✔ Изменения сохранены');
  } catch (err) {
    showToast('✖ Не удалось сохранить файл: ' + err.message, 'error');
  }
}

// ============================================================
//  Save As — write to a new file; new file becomes current
// ============================================================
async function handleSaveAs() {
  if (!currentFilePath) return;

  const defaultName = buildDefaultSaveAsName();
  const defaultPath = currentFilePath.replace(/[^/\\]+$/, defaultName);

  let targetPath;
  try {
    targetPath = await window.electronAPI.saveFileDialog(defaultPath);
  } catch (err) {
    showToast('✖ Ошибка при открытии диалога: ' + err.message, 'error');
    return;
  }

  if (!targetPath) return; // user cancelled

  const updates = buildUpdates();
  try {
    await window.electronAPI.writeExcel(currentFilePath, targetPath, updates);
    // New file becomes the current working file
    currentFilePath = targetPath;
    filePathDisplay.textContent = targetPath;
    commitCurrentValues();
    showToast('✔ Файл успешно сохранен');
  } catch (err) {
    showToast('✖ Не удалось сохранить файл: ' + err.message, 'error');
  }
}

// ============================================================
//  Close-app flow: main process asks us to save then close
// ============================================================
window.electronAPI.onRequestSaveBeforeClose(async () => {
  await handleSave();
  window.electronAPI.closeApp();
});

// ============================================================
//  Collapsible card logic
// ============================================================
function setCardOpen(cardId, open) {
  const card = document.getElementById('card-' + cardId);
  if (!card) return;
  const btn = card.querySelector('.card-title[data-toggle]');
  card.classList.toggle('is-open', open);
  if (btn) btn.setAttribute('aria-expanded', String(open));
  cardOpenState[cardId] = open;
}

function toggleCard(cardId) {
  setCardOpen(cardId, !cardOpenState[cardId]);
}

document.querySelectorAll('.card-title[data-toggle]').forEach((btn) => {
  btn.addEventListener('click', () => toggleCard(btn.getAttribute('data-toggle')));
});

// ============================================================
//  Visibility pass — hide empty rows and empty cards
// ============================================================
function applyVisibility() {
  const allCardIds = ['deal', 'property', 'seller', 'owner1', 'owner2', 'owner3', 'buyer'];

  allCardIds.forEach((cardId) => {
    const card = document.getElementById('card-' + cardId);
    if (!card) return;

    // Show all rows initially, then hide empty ones
    card.querySelectorAll('.field-row[data-field-row]').forEach((row) => {
      const input = row.querySelector('input[type="text"]');
      row.hidden = input ? isFieldEmpty(input.value) : true;
    });

    if (isSectionEmpty(cardId)) {
      card.hidden = true;
      return;
    }

    card.hidden = false;
    setCardOpen(cardId, AUTO_OPEN_CARDS.has(cardId));
  });
}

// ============================================================
//  Populate form from parsed data object
// ============================================================
function populateForm(data) {
  clearAllInputs();
  dirtyInputIds.clear();
  originalValues = {};

  // Extract and store the row map
  rowMap = data._rowMap || {};

  const blockKeys = ['deal', 'property', 'seller', 'owner1', 'owner2', 'owner3', 'buyer'];
  blockKeys.forEach((block) => {
    const blockData = data[block];
    if (!blockData) return;
    Object.entries(blockData).forEach(([fieldName, value]) => {
      const mapKey  = `${block}-${fieldName}`;
      const inputId = FIELD_MAP[mapKey];
      if (inputId) setInputValue(inputId, value);
    });
  });

  // Snapshot clean baseline
  commitCurrentValues();

  // Enable Save As now that a file is loaded
  btnSaveAs.disabled = false;

  applyVisibility();
}

// ============================================================
//  Main flow — choose and load Excel file
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

  currentFilePath = filePath;
  populateForm(data);

  emptyState.hidden = true;
  cardsGrid.hidden  = false;
}

// ============================================================
//  Event listeners
// ============================================================
btnChooseFile.addEventListener('click', handleChooseFile);
btnSave.addEventListener('click', handleSave);
btnSaveAs.addEventListener('click', handleSaveAs);
errorClose.addEventListener('click', hideError);
