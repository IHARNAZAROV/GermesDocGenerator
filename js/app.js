'use strict';

// ============================================================
//  DOM references
// ============================================================
const btnChooseFile   = document.getElementById('btn-choose-file');
const btnSave         = document.getElementById('btn-save');
const btnSaveAs       = document.getElementById('btn-save-as');
const btnClear        = document.getElementById('btn-clear');
const btnCheck        = document.getElementById('btn-check');
const btnGenerate     = document.getElementById('btn-generate');
const btnPreview      = document.getElementById('btn-preview');
const btnSelectAll    = document.getElementById('btn-select-all');
const btnDeselectAll  = document.getElementById('btn-deselect-all');
const filePathDisplay = document.getElementById('file-path-display');
const fileSuccess     = document.getElementById('file-success');
const fileName        = document.getElementById('file-name');
const statusText      = document.getElementById('status-text');
const errorBanner     = document.getElementById('error-banner');
const errorText       = document.getElementById('error-text');
const errorClose      = document.getElementById('error-close');
const loader          = document.getElementById('loader');
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

// Reverse map: inputId → mapKey (for save lookup)
const SAVE_MAP = Object.fromEntries(
  Object.entries(FIELD_MAP).map(([mapKey, inputId]) => [inputId, mapKey])
);

// ============================================================
//  Application state
// ============================================================
let currentFilePath = null;
let rowMap          = {};
let originalValues  = {};
let dirtyInputIds   = new Set();

// ============================================================
//  Universal helpers
// ============================================================
function isFieldEmpty(value) {
  if (value === null || value === undefined) return true;
  return String(value).trim() === '';
}

// ============================================================
//  Toast notifications
// ============================================================
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
//  Error / loader helpers
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

function setStatus(text) {
  if (statusText) statusText.textContent = text;
}

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
  setStatus(hasDirty ? 'Есть несохранённые изменения' : (currentFilePath ? 'Файл загружен' : 'Готов к работе'));
}

function onInputChange(inputId, currentValue) {
  const original = originalValues[inputId] ?? '';
  const current  = currentValue.trim();
  const el = document.getElementById(inputId);
  if (!el) return;
  if (current !== original) {
    dirtyInputIds.add(inputId);
    el.classList.add('input-dirty');
  } else {
    dirtyInputIds.delete(inputId);
    el.classList.remove('input-dirty');
  }
  updateDirtyState();
  // Re-evaluate contract availability whenever any owner field changes
  if (inputId.startsWith('owner1-') || inputId.startsWith('owner2-') || inputId.startsWith('owner3-')) {
    updateContractAvailability();
  }
}

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

// Attach change listeners to every tracked input at startup
for (const inputId of Object.values(FIELD_MAP)) {
  const el = document.getElementById(inputId);
  if (!el) continue;
  el.addEventListener('input', () => onInputChange(inputId, el.value));
}

// ============================================================
//  Build updates map for writing to Excel
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
//  Default "Save As" filename
// ============================================================
function buildDefaultSaveAsName() {
  const lastName = (document.getElementById('seller-Фамилия')?.value || '').trim();
  const dealDate = (document.getElementById('deal-Дата договора')?.value || '').trim();
  let dateStr = dealDate;
  const dateParts = dealDate.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (dateParts) dateStr = `${dateParts[3]}-${dateParts[2]}-${dateParts[1]}`;
  const parts = ['Сделка'];
  if (lastName) parts.push(lastName);
  if (dateStr)  parts.push(dateStr);
  return parts.join('_') + '.xlsx';
}

// ============================================================
//  Save
// ============================================================
async function handleSave() {
  if (!currentFilePath || dirtyInputIds.size === 0) return;
  const updates = buildUpdates();
  try {
    await window.electronAPI.writeExcel(currentFilePath, currentFilePath, updates);
    commitCurrentValues();
    setStatus('Изменения сохранены');
    showToast('✔ Изменения сохранены');
  } catch (err) {
    showToast('✖ Не удалось сохранить файл: ' + err.message, 'error');
  }
}

// ============================================================
//  Save As
// ============================================================
async function handleSaveAs() {
  if (!currentFilePath) return;
  const defaultPath = currentFilePath.replace(/[^/\\]+$/, buildDefaultSaveAsName());
  let targetPath;
  try {
    targetPath = await window.electronAPI.saveFileDialog(defaultPath);
  } catch (err) {
    showToast('✖ Ошибка при открытии диалога: ' + err.message, 'error');
    return;
  }
  if (!targetPath) return;
  const updates = buildUpdates();
  try {
    await window.electronAPI.writeExcel(currentFilePath, targetPath, updates);
    currentFilePath = targetPath;
    filePathDisplay.value = targetPath;
    const baseName = targetPath.split(/[\\/]/).pop();
    fileName.textContent = baseName;
    commitCurrentValues();
    setStatus('Файл сохранён: ' + baseName);
    showToast('✔ Файл успешно сохранён');
  } catch (err) {
    showToast('✖ Не удалось сохранить файл: ' + err.message, 'error');
  }
}

// ============================================================
//  Close-app: main process asks renderer to save-then-close
// ============================================================
window.electronAPI.onRequestSaveBeforeClose(async () => {
  await handleSave();
  window.electronAPI.closeApp();
});

// ============================================================
//  Owner tabs
// ============================================================
const tabBtns  = document.querySelectorAll('.tab-btn[data-tab]');
const tabPanes = document.querySelectorAll('.tab-pane');

function switchTab(tabId) {
  tabBtns.forEach((btn) => btn.classList.toggle('active', btn.dataset.tab === tabId));
  tabPanes.forEach((pane) => pane.classList.toggle('active', pane.id === 'tab-pane-' + tabId));
}

tabBtns.forEach((btn) => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});

// ============================================================
//  Populate form from parsed data
// ============================================================
function populateForm(data) {
  clearAllInputs();
  dirtyInputIds.clear();
  originalValues = {};

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

  // Always recompute "Цена прописью" from "Цена BYN" — ignore any value from Excel
  autoUpdatePropis();

  commitCurrentValues();
  btnSaveAs.disabled = false;
  switchTab('owner1');
  updateContractAvailability();
}

// ============================================================
//  Clear form
// ============================================================
function handleClearForm() {
  clearAllInputs();
  rowMap = {};
  originalValues = {};
  dirtyInputIds.clear();
  currentFilePath = null;
  filePathDisplay.value = '';
  fileSuccess.hidden = true;
  btnSave.disabled = true;
  btnSaveAs.disabled = true;
  setStatus('Готов к работе');
  switchTab('owner1');
  resetContractAvailability();
}

// ============================================================
//  Check data (basic validation)
// ============================================================
function handleCheckData() {
  const missing = [];
  const required = [
    ['deal-Дата договора',      'Дата договора'],
    ['deal-Тип договора',       'Тип договора'],
    ['property-Адрес',          'Адрес объекта'],
    ['seller-Фамилия',          'Фамилия продавца'],
    ['buyer-Фамилия',           'Фамилия покупателя'],
  ];
  required.forEach(([id, label]) => {
    const el = document.getElementById(id);
    if (el && isFieldEmpty(el.value)) missing.push(label);
  });
  if (missing.length === 0) {
    showToast('✔ Основные поля заполнены');
  } else {
    showToast('✖ Не заполнены: ' + missing.join(', '), 'error');
  }
}

// ============================================================
//  Auto-compute "Цена прописью" from "Цена BYN"
// ============================================================
const bynInput    = document.getElementById('deal-Стоимость BYN');
const propisInput = document.getElementById('deal-Стоимость прописью');
const bynErrorEl  = document.getElementById('byn-error');

function autoUpdatePropis() {
  if (!bynInput || !propisInput) return;

  // Normalise: treat comma as decimal separator
  const raw = bynInput.value.replace(',', '.').trim();

  if (raw === '') {
    propisInput.value = '';
    if (bynErrorEl) bynErrorEl.hidden = true;
    bynInput.classList.remove('byn-input-error');
    return;
  }

  // Validate: digits only, optional single dot, max 2 decimal places, no negatives
  const valid = /^\d+(\.\d{1,2})?$/.test(raw) && parseFloat(raw) >= 0;

  if (!valid) {
    if (bynErrorEl) bynErrorEl.hidden = false;
    bynInput.classList.add('byn-input-error');
    propisInput.value = '';
    return;
  }

  if (bynErrorEl) bynErrorEl.hidden = true;
  bynInput.classList.remove('byn-input-error');
  propisInput.value = window.moneyToText(raw);
}

if (bynInput) {
  bynInput.addEventListener('input', autoUpdatePropis);
}

// ============================================================
//  Owners count detection
// ============================================================
const OWNER_SIGNIFICANT_FIELDS = [
  'Фамилия', 'Имя', 'Паспорт серия', 'Паспорт номер', 'Идентификационный номер',
];

function isOwnerPresent(ownerPrefix) {
  return OWNER_SIGNIFICANT_FIELDS.some((field) => {
    const el = document.getElementById(`${ownerPrefix}-${field}`);
    return el && !isFieldEmpty(el.value);
  });
}

function getOwnersCount() {
  if (isOwnerPresent('owner3')) return 3;
  if (isOwnerPresent('owner2')) return 2;
  return 1;
}

// ============================================================
//  Contract availability
// ============================================================
const OWNERS_TOOLTIP = {
  1: 'Недоступно. В сделке участвует только один собственник.',
  2: 'Недоступно. В сделке участвуют два собственника.',
  3: 'Недоступно. В сделке участвуют три собственника.',
};

function updateContractAvailability() {
  const count    = getOwnersCount();
  const tooltip  = OWNERS_TOOLTIP[count] || '';
  document.querySelectorAll('.tpl-item[data-owners-required]').forEach((label) => {
    const required   = label.dataset.ownersRequired;
    const cb         = label.querySelector('input[type="checkbox"]');
    const isDisabled = required !== 'any' && Number(required) !== count;
    if (isDisabled) {
      label.classList.add('tpl-item-disabled');
      label.dataset.tooltip = tooltip;
      if (cb) { cb.disabled = true; cb.checked = false; }
    } else {
      label.classList.remove('tpl-item-disabled');
      delete label.dataset.tooltip;
      if (cb) cb.disabled = false;
    }
  });
}

function resetContractAvailability() {
  document.querySelectorAll('.tpl-item[data-owners-required]').forEach((label) => {
    label.classList.remove('tpl-item-disabled');
    delete label.dataset.tooltip;
    const cb = label.querySelector('input[type="checkbox"]');
    if (cb) cb.disabled = false;
  });
}

// ============================================================
//  Template checkboxes
// ============================================================
function handleSelectAll() {
  document.querySelectorAll('.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]').forEach((cb) => { cb.checked = true; });
}
function handleDeselectAll() {
  document.querySelectorAll('.tpl-item input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
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
  if (!filePath) return;

  filePathDisplay.value = filePath;
  fileSuccess.hidden = true;
  setStatus('Чтение файла…');
  showLoader();

  let data;
  try {
    data = await window.electronAPI.readExcel(filePath);
  } catch (err) {
    hideLoader();
    showError('Ошибка при чтении файла: ' + err.message);
    setStatus('Ошибка чтения файла');
    return;
  }

  hideLoader();

  if (!data || typeof data !== 'object') {
    showError('Файл прочитан, но данные не получены. Проверьте формат файла.');
    setStatus('Ошибка формата файла');
    return;
  }

  currentFilePath = filePath;
  const baseName = filePath.split(/[\\/]/).pop();
  fileName.textContent = baseName;
  fileSuccess.hidden = false;

  populateForm(data);
  setStatus('Файл загружен: ' + baseName);
}

// ============================================================
//  Event listeners
// ============================================================
btnChooseFile.addEventListener('click', handleChooseFile);
btnSave.addEventListener('click', handleSave);
btnSaveAs.addEventListener('click', handleSaveAs);
btnClear.addEventListener('click', handleClearForm);
btnCheck.addEventListener('click', handleCheckData);
errorClose.addEventListener('click', hideError);
btnSelectAll.addEventListener('click', handleSelectAll);
btnDeselectAll.addEventListener('click', handleDeselectAll);

btnGenerate.addEventListener('click', () => {
  showToast('Функция генерации договоров будет доступна в следующей версии');
});
btnPreview.addEventListener('click', () => {
  showToast('Предварительный просмотр будет доступен в следующей версии');
});
