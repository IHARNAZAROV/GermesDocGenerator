'use strict';

// ============================================================
//  Динамическая генерация формы из FIELDS_CONFIG
//  FIELD_MAP строится автоматически — не редактируйте вручную.
//  Чтобы добавить/удалить поля:
//    1. Обновите шаблон Excel
//    2. Запустите: node scripts/scan-excel.js <путь/к/шаблону.xlsx>
//    3. Перезапустите приложение
// ============================================================
const FIELD_MAP = window.FormBuilder.buildForm(window.FIELDS_CONFIG);

// Reverse map: inputId → mapKey (для поиска при сохранении)
const SAVE_MAP = Object.fromEntries(
  Object.entries(FIELD_MAP).map(([mapKey, inputId]) => [inputId, mapKey])
);

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
const btnBrowse       = document.getElementById('btn-browse');
const saveFolderInput = document.getElementById('save-folder');
const chkOpenAfter    = document.getElementById('chk-open-after');
const chkAddDate      = document.getElementById('chk-add-date');
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
function showToast(message, type = 'success', duration = 3200) {
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
  }, duration);
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
  // Re-evaluate contract availability whenever owner fields or seller-ownership flag changes
  if (
    inputId.startsWith('owner1-') ||
    inputId.startsWith('owner2-') ||
    inputId.startsWith('owner3-') ||
    inputId === 'seller-Является собственником'
  ) {
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

// Event delegation: один слушатель на общий предок вместо N слушателей на каждый input.
// FIELD_IDS — Set для O(1)-проверки принадлежности элемента к отслеживаемым полям.
const FIELD_IDS = new Set(Object.values(FIELD_MAP));

document.getElementById('deal-body').addEventListener('input', (e) => {
  const id = e.target.id;
  if (!id || !FIELD_IDS.has(id)) return;
  onInputChange(id, e.target.value);
});

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
  // Always recompute commission — ignore any value from Excel
  autoUpdateCommission();

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
  if (commissionInput) commissionInput.value = '';
  setStatus('Готов к работе');
  switchTab('owner1');
  resetContractAvailability();
}

// ============================================================
//  Check data (basic validation)
// ============================================================
function handleCheckData() {
  const missing = [];

  function chk(id, label) {
    const el = document.getElementById(id);
    if (el && isFieldEmpty(el.value)) missing.push(label);
  }

  // ── СДЕЛКА ───────────────────────────────────────────────
  chk('deal-Стоимость BYN',              'Сделка → Цена BYN');
  chk('deal-Номер сделки',               'Сделка → Номер сделки');
  chk('deal-Дата договора',              'Сделка → Дата договора');
  chk('deal-Дата окончания договора',    'Сделка → Дата окончания договора');
  chk('deal-Тип договора',              'Сделка → Тип договора');
  chk('deal-Количество собственников',  'Сделка → Кол-во собственников');
  chk('deal-Ответственный риэлтер',     'Сделка → Риэлтер');
  // Комиссия и Стоимость прописью — вычисляются, не проверяем

  // ── ОБЪЕКТ ───────────────────────────────────────────────
  chk('property-Тип объекта',    'Объект → Тип объекта');
  chk('property-Адрес',          'Объект → Адрес');
  chk('property-Город',          'Объект → Город');
  chk('property-Улица',          'Объект → Улица');
  chk('property-Дом',            'Объект → Дом');
  chk('property-Этаж',           'Объект → Этаж');
  chk('property-Этажность',      'Объект → Этажность');
  chk('property-Количество комнат', 'Объект → Кол-во комнат');
  chk('property-Общая площадь',  'Объект → Общая площадь');
  chk('property-Жилая площадь',  'Объект → Жилая площадь');
  chk('property-Площадь кухни',  'Объект → Площадь кухни');
  chk('property-Кадастровый номер',  'Объект → Кадастровый №');
  chk('property-Инвентарный номер',  'Объект → Инвентарный №');

  // ── ПРОДАВЕЦ ─────────────────────────────────────────────
  chk('seller-Фамилия',                 'Продавец → Фамилия');
  chk('seller-Имя',                     'Продавец → Имя');
  chk('seller-Отчество',                'Продавец → Отчество');
  chk('seller-Дата рождения',           'Продавец → Дата рождения');
  chk('seller-Паспорт серия',           'Продавец → Паспорт серия');
  chk('seller-Паспорт номер',           'Продавец → Паспорт номер');
  chk('seller-Идентификационный номер', 'Продавец → Идент. номер');
  chk('seller-Кем выдан',              'Продавец → Кем выдан');
  chk('seller-Дата выдачи',            'Продавец → Дата выдачи');
  chk('seller-Адрес регистрации',       'Продавец → Адрес регистрации');
  chk('seller-Является собственником',  'Продавец → Является собственником');

  // ── СОБСТВЕННИК №1 ───────────────────────────────────────
  // Проверяем всегда, если продавец не является собственником,
  // либо если в вкладке собственника есть хотя бы фамилия
  if (isOwnerPresent('owner1') || (() => {
    const isOwnerRaw = (getField('seller-Является собственником') || '').trim().toLowerCase();
    return isOwnerRaw === 'нет';
  })()) {
    chk('owner1-Фамилия',                 'Собственник 1 → Фамилия');
    chk('owner1-Имя',                     'Собственник 1 → Имя');
    chk('owner1-Отчество',                'Собственник 1 → Отчество');
    chk('owner1-Дата рождения',           'Собственник 1 → Дата рождения');
    chk('owner1-Паспорт серия',           'Собственник 1 → Паспорт серия');
    chk('owner1-Паспорт номер',           'Собственник 1 → Паспорт номер');
    chk('owner1-Идентификационный номер', 'Собственник 1 → Идент. номер');
    chk('owner1-Кем выдан',              'Собственник 1 → Кем выдан');
    chk('owner1-Дата выдачи',            'Собственник 1 → Дата выдачи');
    chk('owner1-Адрес регистрации',       'Собственник 1 → Адрес регистрации');
    chk('owner1-Доля собственности',      'Собственник 1 → Доля собственности');
  }

  // ── СОБСТВЕННИК №2 (только если присутствует) ────────────
  if (isOwnerPresent('owner2')) {
    chk('owner2-Фамилия',                 'Собственник 2 → Фамилия');
    chk('owner2-Имя',                     'Собственник 2 → Имя');
    chk('owner2-Отчество',                'Собственник 2 → Отчество');
    chk('owner2-Дата рождения',           'Собственник 2 → Дата рождения');
    chk('owner2-Паспорт серия',           'Собственник 2 → Паспорт серия');
    chk('owner2-Паспорт номер',           'Собственник 2 → Паспорт номер');
    chk('owner2-Идентификационный номер', 'Собственник 2 → Идент. номер');
    chk('owner2-Кем выдан',              'Собственник 2 → Кем выдан');
    chk('owner2-Дата выдачи',            'Собственник 2 → Дата выдачи');
    chk('owner2-Адрес регистрации',       'Собственник 2 → Адрес регистрации');
    chk('owner2-Доля собственности',      'Собственник 2 → Доля собственности');
  }

  // ── СОБСТВЕННИК №3 (только если присутствует) ────────────
  if (isOwnerPresent('owner3')) {
    chk('owner3-Фамилия',                 'Собственник 3 → Фамилия');
    chk('owner3-Имя',                     'Собственник 3 → Имя');
    chk('owner3-Отчество',                'Собственник 3 → Отчество');
    chk('owner3-Дата рождения',           'Собственник 3 → Дата рождения');
    chk('owner3-Паспорт серия',           'Собственник 3 → Паспорт серия');
    chk('owner3-Паспорт номер',           'Собственник 3 → Паспорт номер');
    chk('owner3-Идентификационный номер', 'Собственник 3 → Идент. номер');
    chk('owner3-Кем выдан',              'Собственник 3 → Кем выдан');
    chk('owner3-Дата выдачи',            'Собственник 3 → Дата выдачи');
    chk('owner3-Адрес регистрации',       'Собственник 3 → Адрес регистрации');
    chk('owner3-Доля собственности',      'Собственник 3 → Доля собственности');
  }

  // ── ПОКУПАТЕЛЬ ───────────────────────────────────────────
  chk('buyer-Фамилия',                 'Покупатель → Фамилия');
  chk('buyer-Имя',                     'Покупатель → Имя');
  chk('buyer-Отчество',                'Покупатель → Отчество');
  chk('buyer-Дата рождения',           'Покупатель → Дата рождения');
  chk('buyer-Паспорт серия',           'Покупатель → Паспорт серия');
  chk('buyer-Паспорт номер',           'Покупатель → Паспорт номер');
  chk('buyer-Идентификационный номер', 'Покупатель → Идент. номер');
  chk('buyer-Кем выдан',              'Покупатель → Кем выдан');
  chk('buyer-Дата выдачи',            'Покупатель → Дата выдачи');
  chk('buyer-Адрес регистрации',       'Покупатель → Адрес регистрации');

  // ── Результат ────────────────────────────────────────────
  if (missing.length === 0) {
    showToast('✔ Все поля заполнены');
  } else {
    showToast(`✖ Не заполнено полей: ${missing.length}. ${missing.join(' | ')}`, 'error', 6000);
  }
}

// ============================================================
//  Auto-compute "Цена прописью" from "Цена BYN"
// ============================================================
const bynInput         = document.getElementById('deal-Стоимость BYN');
const propisInput      = document.getElementById('deal-Стоимость прописью');
const bynErrorEl       = document.getElementById('byn-error');
const commissionInput  = document.getElementById('deal-Комиссия агентства');

// Make commission field read-only — it is always computed, never entered manually
if (commissionInput) {
  commissionInput.readOnly = true;
  commissionInput.title    = 'Вычисляется автоматически по тарифной таблице';
  commissionInput.style.background = '#f4f6f8';
  commissionInput.style.color      = '#555';
  commissionInput.style.cursor     = 'default';
}

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
//  Auto-compute commission from "Цена BYN"
// ============================================================
function autoUpdateCommission() {
  if (!commissionInput) return;

  const raw = (bynInput ? bynInput.value : '').replace(',', '.').trim();
  if (raw === '' || !/^\d+(\.\d{1,2})?$/.test(raw)) {
    commissionInput.value = '';
    return;
  }

  const result = window.calculateCommission(parseFloat(raw), window.COMMISSION_CONFIG.baseValue);
  commissionInput.value = result.amountBYN ? `${result.amountBYN} (${result.percent}%)` : '';
}

if (bynInput) {
  bynInput.addEventListener('input', autoUpdateCommission);
}

// ============================================================
//  Instrumental case converter for passport-issuing organizations
//  (творительный падеж: «выдан Лидским РОВД» вместо «выдан Лидский РОВД»)
// ============================================================
function toInstrumental(str) {
  if (!str) return str;
  // Replace adjective nominative endings with instrumental equivalents.
  // Lookahead (?=\s|$) targets word-final positions in Cyrillic text.
  // Order matters: longer/more specific endings first.
  return str
    .replace(/(ский)(?=\s|$)/g,  'ским')
    .replace(/(цкий)(?=\s|$)/g,  'цким')
    .replace(/(жний)(?=\s|$)/g,  'жним')
    .replace(/(дний)(?=\s|$)/g,  'дним')
    .replace(/(зний)(?=\s|$)/g,  'зним')
    .replace(/(ний)(?=\s|$)/g,   'ним')
    .replace(/(жный)(?=\s|$)/g,  'жным')
    .replace(/(дный)(?=\s|$)/g,  'дным')
    .replace(/(зный)(?=\s|$)/g,  'зным')
    .replace(/(ный)(?=\s|$)/g,   'ным')
    .replace(/(ий)(?=\s|$)/g,    'им')
    .replace(/(ый)(?=\s|$)/g,    'ым');
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

function getSellerIsOwner() {
  const el = document.getElementById('seller-Является собственником');
  if (!el) return false;
  const raw = el.value.trim().toLowerCase();
  return raw === 'да' || raw === 'yes';
}

function getOwnersCount() {
  if (getSellerIsOwner()) {
    // Продавец — собственник №1; дополнительные совладельцы — во вкладках собственников
    const coOwners = (isOwnerPresent('owner1') ? 1 : 0)
                   + (isOwnerPresent('owner2') ? 1 : 0)
                   + (isOwnerPresent('owner3') ? 1 : 0);
    return 1 + coOwners;
  } else {
    // Продавец действует по доверенности; собственники — только из вкладок
    if (isOwnerPresent('owner3')) return 3;
    if (isOwnerPresent('owner2')) return 2;
    return 1;
  }
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
//  Main flow — load Excel file (shared between dialog & drag-drop)
// ============================================================
const dropZone     = document.getElementById('drop-zone');
const dropIdle     = document.getElementById('drop-idle');
const dropSuccess  = document.getElementById('drop-success');
const dropFileName = document.getElementById('file-name');

function setDropState(state) {
  // state: 'idle' | 'over' | 'success'
  dropIdle.classList.toggle('dz-hidden',    state !== 'idle');
  dropSuccess.classList.toggle('dz-hidden', state !== 'success');
  dropZone.classList.toggle('drop-zone--over',   state === 'over');
  dropZone.classList.toggle('drop-zone--loaded', state === 'success');
}

async function loadExcelFile(filePath) {
  filePathDisplay.value = filePath;
  fileSuccess.hidden = true;
  setStatus('Чтение файла…');
  showLoader();

  let data;
  try {
    data = await window.electronAPI.readExcel(filePath);
  } catch (err) {
    hideLoader();
    setDropState('idle');
    showError('Ошибка при чтении файла: ' + err.message);
    setStatus('Ошибка чтения файла');
    return;
  }

  hideLoader();

  if (!data || typeof data !== 'object') {
    setDropState('idle');
    showError('Файл прочитан, но данные не получены. Проверьте формат файла.');
    setStatus('Ошибка формата файла');
    return;
  }

  currentFilePath = filePath;
  const baseName = filePath.split(/[\\/]/).pop();
  dropFileName.textContent = baseName;
  setDropState('success');

  populateForm(data);
  setStatus('Файл загружен: ' + baseName);
}

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
  await loadExcelFile(filePath);
}

// ============================================================
//  Drag-and-drop handlers
// ============================================================
let dragCounter = 0; // track nested drag-enter/leave

dropZone.addEventListener('dragenter', e => {
  e.preventDefault();
  dragCounter++;
  if (dragCounter === 1) setDropState('over');
});

dropZone.addEventListener('dragleave', () => {
  dragCounter--;
  if (dragCounter === 0) setDropState(currentFilePath ? 'success' : 'idle');
});

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
});

dropZone.addEventListener('drop', async e => {
  e.preventDefault();
  dragCounter = 0;
  hideError();

  const file = e.dataTransfer.files[0];
  if (!file) { setDropState(currentFilePath ? 'success' : 'idle'); return; }

  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    setDropState(currentFilePath ? 'success' : 'idle');
    showError('Поддерживаются только файлы Excel (.xlsx, .xls)');
    return;
  }

  // Electron 32+: webUtils.getPathForFile() через context bridge
  const filePath = window.electronAPI.getPathForFile(file);
  if (!filePath) {
    setDropState(currentFilePath ? 'success' : 'idle');
    showError('Не удалось получить путь к файлу.');
    return;
  }

  await loadExcelFile(filePath);
});

document.getElementById('btn-drop-browse').addEventListener('click', handleChooseFile);
document.getElementById('btn-drop-change').addEventListener('click', handleChooseFile);

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

// ============================================================
//  Scan template — кнопка «Обновить шаблон»
// ============================================================
const btnScanTemplate = document.getElementById('btn-scan-template');
if (btnScanTemplate) {
  btnScanTemplate.addEventListener('click', async () => {
    btnScanTemplate.disabled = true;
    setStatus('Сканирование шаблона…');
    try {
      const result = await window.electronAPI.scanTemplate();
      if (result.canceled) {
        setStatus(currentFilePath ? 'Файл загружен' : 'Готов к работе');
        return;
      }
      if (!result.ok) {
        showToast('✖ Ошибка сканирования: ' + result.error, 'error');
        setStatus('Ошибка сканирования');
        return;
      }
      // main.js перезагружает окно — toast не успеет показаться,
      // но если reload по какой-то причине не случился — покажем сообщение.
      const info = [];
      if (result.added?.length)   info.push(`+${result.added.length} новых`);
      if (result.removed?.length) info.push(`−${result.removed.length} удалено`);
      showToast('✔ Шаблон обновлён' + (info.length ? ': ' + info.join(', ') : ''));
    } catch (err) {
      showToast('✖ ' + err.message, 'error');
      setStatus('Ошибка');
    } finally {
      btnScanTemplate.disabled = false;
    }
  });
}

// ============================================================
//  Browse output folder
// ============================================================
btnBrowse.addEventListener('click', async () => {
  const current = saveFolderInput.value.trim() || undefined;
  const chosen = await window.electronAPI.selectFolder(current);
  if (chosen) saveFolderInput.value = chosen;
});

// ============================================================
//  Template data helpers
// ============================================================
function getField(id) {
  return (document.getElementById(id)?.value || '').trim();
}

// ============================================================
//  Template registry
//  key matches data-template attribute on .tpl-item labels.
//  Each entry: { label, generate(outputDir) → Promise<{success, path?, error?}> }
//
//  Добавление нового шаблона Word:
//    1. Добавьте файл .docx в templates/working/
//    2. Добавьте data-template="your-key" в label в index.html
//    3. Добавьте запись здесь с соответствующим ключом
// ============================================================
// ============================================================
//  buildPlaceholderData() — единый источник данных для всех
//  Word-шаблонов. Структура соответствует config/placeholders.json.
//  Плейсхолдеры в .docx: {{deal.number}}, {{seller.fullName}} и т.д.
// ============================================================
// ============================================================
//  Date → "16 июля 2027" (long Russian format)
// ============================================================
const MONTHS_GEN_RU = [
  'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря',
];

function dateToLongRussian(ddmmyyyy) {
  if (!ddmmyyyy) return '';
  const m = String(ddmmyyyy).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) return ddmmyyyy;
  const day   = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  const year  = m[3];
  if (month < 1 || month > 12) return ddmmyyyy;
  return `${day} ${MONTHS_GEN_RU[month - 1]} ${year}`;
}

const GENITIVE_MAP = {
  'квартира':          'квартиры',
  'дом':               'дома',
  'жилой дом':         'жилого дома',
  'комната':           'комнаты',
  'апартаменты':       'апартаментов',
  'гараж':             'гаража',
  'земельный участок': 'земельного участка',
  'офис':              'офиса',
  'нежилое помещение': 'нежилого помещения',
  'помещение':         'помещения',
};

// ============================================================
//  Dative case (дательный падеж) for Russian/Belarusian names
//  «квартира принадлежит КОМУ?» — Малейкович Юзефе Казимировне
//  Gender is auto-detected from the patronymic ending.
// ============================================================
function _declinePatronymicDative(middle) {
  if (!middle) return middle;
  const t = middle.trim();
  const low = t.toLowerCase();
  if (low.endsWith('иична'))  return t.slice(0, -1) + 'е'; // Ильинична → Ильиничне
  if (low.endsWith('овна'))   return t.slice(0, -1) + 'е'; // Казимировна → Казимировне
  if (low.endsWith('евна'))   return t.slice(0, -1) + 'е'; // Андреевна → Андреевне
  if (low.endsWith('ична'))   return t.slice(0, -1) + 'е'; // Ильична → Ильичне
  if (low.endsWith('ович'))   return t + 'у';              // Иванович → Ивановичу
  if (low.endsWith('евич'))   return t + 'у';              // Андреевич → Андреевичу
  if (low.endsWith('ич'))     return t + 'у';              // Ильич → Ильичу
  return t;
}

function _declineFirstNameDative(first, gender) {
  if (!first) return first;
  const t = first.trim();
  const low = t.toLowerCase();
  if (gender === 'f' || gender === null) {
    if (low.endsWith('ия'))  return t.slice(0, -1) + 'и'; // Мария → Марии
    if (low.endsWith('ья'))  return t.slice(0, -1) + 'е'; // Дарья/Наталья → Дарье/Наталье
    if (low.endsWith('я'))   return t.slice(0, -1) + 'е'; // Таня → Тане
    if (low.endsWith('а'))   return t.slice(0, -1) + 'е'; // Юзефа → Юзефе, Анна → Анне
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'и'; // Любовь → Любови
  }
  if (gender === 'm') {
    if (low.endsWith('ий'))  return t.slice(0, -2) + 'ию'; // Василий → Василию
    if (low.endsWith('й'))   return t.slice(0, -1) + 'ю';  // Сергей → Сергею
    if (low.endsWith('я'))   return t.slice(0, -1) + 'е';  // Илья → Илье
    if (low.endsWith('а'))   return t.slice(0, -1) + 'е';  // Никита → Никите
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'ю';  // Игорь → Игорю
    return t + 'у'; // Иван → Ивану, Александр → Александру
  }
  return t;
}

function _declineLastNameDative(last, gender) {
  if (!last) return last;
  const t = last.trim();
  const low = t.toLowerCase();
  if (gender === 'f') {
    if (low.endsWith('ская') || low.endsWith('цкая') || low.endsWith('зская')) {
      return t.slice(0, -2) + 'ой';  // Островская → Островской
    }
    if (low.endsWith('ова') || low.endsWith('ева') || low.endsWith('ёва') ||
        low.endsWith('ина') || low.endsWith('ына')) {
      return t.slice(0, -1) + 'ой';  // Иванова → Ивановой, Пушкина → Пушкиной
    }
    // Ends in consonant (e.g. Малейкович) — feminine surnames don't decline
    return t;
  }
  if (gender === 'm') {
    if (low.endsWith('ский') || low.endsWith('цкий') || low.endsWith('зский')) {
      return t.slice(0, -2) + 'ому'; // Троцкий → Троцкому
    }
    if (low.endsWith('ой') || low.endsWith('ый') || low.endsWith('ий')) {
      return t.slice(0, -2) + 'ому'; // Толстой → Толстому
    }
    if (low.endsWith('ов') || low.endsWith('ев') || low.endsWith('ёв')) {
      return t + 'у'; // Иванов → Иванову
    }
    if (low.endsWith('ин') || low.endsWith('ын')) {
      return t + 'у'; // Пушкин → Пушкину
    }
    if (low.endsWith('ь')) {
      return t.slice(0, -1) + 'ю'; // Медведь → Медведю
    }
    return t + 'у'; // прочие мужские на согласный
  }
  return t;
}

// ============================================================
//  Genitive case (родительный падеж) for Russian/Belarusian names
//  «справка о здоровье КОГО?» — Малейкович Юзефы Казимировны
//  Gender is auto-detected from the patronymic ending.
// ============================================================
function _declinePatronymicGenitive(middle) {
  if (!middle) return middle;
  const t = middle.trim();
  const low = t.toLowerCase();
  if (low.endsWith('иична'))  return t.slice(0, -1) + 'ы'; // Ильиична → Ильиичны
  if (low.endsWith('овна'))   return t.slice(0, -1) + 'ы'; // Казимировна → Казимировны
  if (low.endsWith('евна'))   return t.slice(0, -1) + 'ы'; // Андреевна → Андреевны
  if (low.endsWith('ична'))   return t.slice(0, -1) + 'ы'; // Ильична → Ильичны
  if (low.endsWith('ович'))   return t + 'а';              // Иванович → Ивановича
  if (low.endsWith('евич'))   return t + 'а';              // Андреевич → Андреевича
  if (low.endsWith('ич'))     return t + 'а';              // Ильич → Ильича
  return t;
}

function _declineFirstNameGenitive(first, gender) {
  if (!first) return first;
  const t = first.trim();
  const low = t.toLowerCase();
  const velars = ['г','к','х','ж','ш','щ','ч'];
  if (gender === 'f' || gender === null) {
    if (low.endsWith('ия'))  return t.slice(0, -1) + 'и'; // Мария → Марии
    if (low.endsWith('ья'))  return t.slice(0, -2) + 'ьи'; // Дарья → Дарьи
    if (low.endsWith('я'))   return t.slice(0, -1) + 'и'; // Таня → Тани
    if (low.endsWith('а'))   return velars.includes(low[low.length - 2])
      ? t.slice(0, -1) + 'и'   // Ольга → Ольги
      : t.slice(0, -1) + 'ы';  // Юзефа → Юзефы, Анна → Анны
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'и'; // Любовь → Любови
  }
  if (gender === 'm') {
    if (low.endsWith('ий'))  return t.slice(0, -2) + 'ия'; // Василий → Василия
    if (low.endsWith('й'))   return t.slice(0, -1) + 'я';  // Сергей → Сергея
    if (low.endsWith('я'))   return t.slice(0, -1) + 'и';  // Илья → Ильи
    if (low.endsWith('а'))   return velars.includes(low[low.length - 2])
      ? t.slice(0, -1) + 'и'   // (редко, но по правилу)
      : t.slice(0, -1) + 'ы';  // Никита → Никиты
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'я';  // Игорь → Игоря
    return t + 'а'; // Иван → Ивана, Александр → Александра
  }
  return t;
}

function _declineLastNameGenitive(last, gender) {
  if (!last) return last;
  const t = last.trim();
  const low = t.toLowerCase();
  if (gender === 'f') {
    if (low.endsWith('ская') || low.endsWith('цкая') || low.endsWith('зская')) {
      return t.slice(0, -2) + 'ой';  // Островская → Островской (= dative)
    }
    if (low.endsWith('ова') || low.endsWith('ева') || low.endsWith('ёва') ||
        low.endsWith('ина') || low.endsWith('ына')) {
      return t.slice(0, -1) + 'ой';  // Иванова → Ивановой (= dative)
    }
    // Фамилия на согласный — не склоняется (Малейкович → Малейкович)
    return t;
  }
  if (gender === 'm') {
    if (low.endsWith('ский') || low.endsWith('цкий') || low.endsWith('зский')) {
      return t.slice(0, -2) + 'ого'; // Троцкий → Троцкого
    }
    if (low.endsWith('ой') || low.endsWith('ый') || low.endsWith('ий')) {
      return t.slice(0, -2) + 'ого'; // Толстой → Толстого
    }
    if (low.endsWith('ов') || low.endsWith('ев') || low.endsWith('ёв')) {
      return t + 'а'; // Иванов → Иванова
    }
    if (low.endsWith('ин') || low.endsWith('ын')) {
      return t + 'а'; // Пушкин → Пушкина
    }
    if (low.endsWith('ь')) {
      return t.slice(0, -1) + 'я'; // Медведь → Медведя
    }
    return t + 'а'; // прочие мужские на согласный: Горбач → Горбача
  }
  return t;
}

function buildNameGenitive(lastName, firstName, middleName) {
  const p = (middleName || '').trim().toLowerCase();
  let gender = null;
  if (p.endsWith('иична') || p.endsWith('овна') || p.endsWith('евна') || p.endsWith('ична')) {
    gender = 'f';
  } else if (p.endsWith('ович') || p.endsWith('евич') || p.endsWith('ич')) {
    gender = 'm';
  }
  const lg = _declineLastNameGenitive(lastName || '', gender);
  const fg = _declineFirstNameGenitive(firstName || '', gender);
  const mg = _declinePatronymicGenitive(middleName || '');
  return {
    lastNameGenitive:   lg,
    firstNameGenitive:  fg,
    middleNameGenitive: mg,
    fullNameGenitive:   [lg, fg, mg].filter(Boolean).join(' '),
  };
}

function buildNameDative(lastName, firstName, middleName) {
  const p = (middleName || '').trim().toLowerCase();
  let gender = null;
  if (p.endsWith('иична') || p.endsWith('овна') || p.endsWith('евна') || p.endsWith('ична')) {
    gender = 'f';
  } else if (p.endsWith('ович') || p.endsWith('евич') || p.endsWith('ич')) {
    gender = 'm';
  }
  const ld = _declineLastNameDative(lastName || '', gender);
  const fd = _declineFirstNameDative(firstName || '', gender);
  const md = _declinePatronymicDative(middleName || '');
  return {
    lastNameDative:   ld,
    firstNameDative:  fd,
    middleNameDative: md,
    fullNameDative:   [ld, fd, md].filter(Boolean).join(' '),
  };
}

function buildPersonBlock(prefix) {
  const lastName   = getField(prefix + 'Фамилия')   || '';
  const firstName  = getField(prefix + 'Имя')        || '';
  const middleName = getField(prefix + 'Отчество')   || '';
  const fullName   = [lastName, firstName, middleName].filter(Boolean).join(' ');
  const initials   = lastName && firstName
    ? lastName + ' ' + firstName[0] + '.' + (middleName ? middleName[0] + '.' : '')
    : fullName;
  const series  = getField(prefix + 'Паспорт серия') || '';
  const number  = getField(prefix + 'Паспорт номер') || '';
  const genitive = buildNameGenitive(lastName, firstName, middleName);
  const dative   = buildNameDative(lastName, firstName, middleName);
  return {
    lastName,
    firstName,
    middleName,
    fullName,
    initials,
    ...genitive,
    ...dative,
    birthDate:                    getField(prefix + 'Дата рождения')          || '',
    passportSeries:               series,
    passportNumber:               number,
    passport:                     [series, number].filter(Boolean).join(' '),
    id:                           getField(prefix + 'Идентификационный номер') || '',
    passportIssuedBy:             getField(prefix + 'Кем выдан')              || '',
    passportIssuedByInstrumental: toInstrumental(getField(prefix + 'Кем выдан') || ''),
    passportIssueDate:            getField(prefix + 'Дата выдачи')            || '',
    address:                      getField(prefix + 'Адрес регистрации')      || '',
    phone:                        getField(prefix + 'Телефон')                || '',
    email:                        '',
  };
}

function buildPlaceholderData() {
  const propertyTypeRaw = (getField('property-Тип объекта') || '').trim().toLowerCase();

  const _endDateRaw = getField('deal-Дата окончания договора') || '';
  const deal = {
    number:                    getField('deal-Номер сделки')    || '',
    date:                      getField('deal-Дата договора')   || '',
    dateText:                  '',
    endDate:                   _endDateRaw,
    endDateText:               dateToLongRussian(_endDateRaw),
    exclusive:                 getField('deal-Эксклюзив')       || '',
    contractNumber:            getField('deal-Номер сделки')    || '',
    contractDate:              getField('deal-Дата договора')   || '',
    advertisingContractNumber: '',
    advertisingContractDate:   '',
    depositContractNumber:     '',
    depositContractDate:       '',
    storageContractNumber:     '',
    storageContractDate:       '',
    btiPayment:                getField('deal-Оплата услуг БТИ') || '',
    furniture:                 getField('deal-Мебель')            || '',
  };

  const property = {
    type:         getField('property-Тип объекта')      || '',
    typeGenitive: GENITIVE_MAP[propertyTypeRaw] || getField('property-Тип объекта') || '',
    city:         getField('property-Город')            || '',
    street:       getField('property-Улица')            || '',
    house:        [getField('property-Дом'), getField('property-Корпус')].filter(Boolean).join('/'),
    flat:         getField('property-Квартира')         || '',
    address:      getField('property-Адрес')            || '',
    rooms:        getField('property-Количество комнат')|| '',
    floor:        getField('property-Этаж')             || '',
    floors:       getField('property-Этажность')        || '',
    areaTotal:    getField('property-Общая площадь')    || '',
    areaLiving:   getField('property-Жилая площадь')   || '',
    areaKitchen:  getField('property-Площадь кухни')   || '',
    cadastre:        getField('property-Кадастровый номер') || '',
    inventoryNumber: getField('property-Инвентарный номер')|| '',
    wallMaterial:    getField('property-Материал стен')    || '',
    yearBuilt:    getField('property-Год постройки')   || '',
    priceUSD:          getField('deal-Стоимость USD')   || '',
    priceBYN:          getField('deal-Стоимость BYN')   || '',
    priceWords:        getField('deal-Стоимость прописью') || '',
    priceWordsUSD:     (() => {
      const raw = (getField('deal-Стоимость USD') || '').replace(',', '.').trim();
      return raw ? window.moneyToTextUSD(raw) : '';
    })(),
    remainderUSD:      '',   // заполняется ниже, после вычисления задатка
    remainderUSDWords: '',   // заполняется ниже
  };

  const owner1 = { ...buildPersonBlock('owner1-'), share: getField('owner1-Доля собственности') || '' };
  const owner2 = { ...buildPersonBlock('owner2-'), share: getField('owner2-Доля собственности') || '' };
  const owner3 = { ...buildPersonBlock('owner3-'), share: getField('owner3-Доля собственности') || '' };
  const buyer  = buildPersonBlock('buyer-');

  // ── Логика продавца ────────────────────────────────────────
  // Продавец всегда заполняется из блока ПРОДАВЕЦ в Excel.
  // Если «Является собственником» = «Да» — продавец является собственником №1,
  // дополнительные совладельцы — во вкладках Собственник 1/2/3.
  // Если «Нет» — продавец действует по доверенности; все собственники — из вкладок.
  const isOwnerRaw    = (getField('seller-Является собственником') || '').trim().toLowerCase();
  const sellerIsOwner = isOwnerRaw === 'да' || isOwnerRaw === 'yes';

  const seller = {
    ...buildPersonBlock('seller-'),
    isOwner:   sellerIsOwner,
    poaNumber: sellerIsOwner ? '' : (getField('seller-Номер доверенности') || ''),
    poaDate:   sellerIsOwner ? '' : (getField('seller-Дата доверенности')  || ''),
  };

  // Look up agent from AGENTS_CONFIG by matching last name (or any part of the name)
  // against the value in Excel cell B10 (field «Ответственный риэлтер»)
  const agentRaw = getField('deal-Ответственный риэлтер') || '';
  const agentNormalized = agentRaw.trim().toLowerCase();
  const agentRecord = (window.AGENTS_CONFIG?.agents || []).find((a) =>
    a.matchKeys.some((key) => agentNormalized.includes(key))
  );
  const agent = agentRecord
    ? {
        lastName:          agentRecord.lastName,
        firstName:         agentRecord.firstName,
        middleName:        agentRecord.middleName,
        fullName:          agentRecord.fullName,
        initials:          agentRecord.initials,
        phone:             agentRecord.phone,
        email:             agentRecord.email,
        attestationNumber: agentRecord.attestationNumber,
        attestationDate:   agentRecord.attestationDate,
        attestationExpiry: agentRecord.attestationExpiry,
        cardNumber:        agentRecord.cardNumber,
        cardDate:          agentRecord.cardDate,
      }
    : {
        lastName:          '',
        firstName:         '',
        middleName:        '',
        fullName:          agentRaw,
        initials:          agentRaw,
        phone:             '',
        email:             '',
        attestationNumber: '',
        attestationDate:   '',
        attestationExpiry: '',
        cardNumber:        '',
        cardDate:          '',
      };

  const agency = {
    name: '', shortName: '', director: '', address: '',
    phone: '', email: '', website: '', bank: '',
    bankAccount: '', bik: '', unp: '',
  };

  const keys  = { count: '', countWords: '' };
  const money = { amount: '', amountWords: '', currency: '' };

  // ── Комиссия агентства ──────────────────────────────────────
  // Вычисляется по тарифной таблице: priceBYN / baseValue → строка → процент → сумма
  const priceBYNRaw = parseFloat(
    (getField('deal-Стоимость BYN') || '0').replace(',', '.')
  );
  const commissionResult = window.calculateCommission(
    priceBYNRaw,
    window.COMMISSION_CONFIG.baseValue
  );
  const commission = {
    percent:     commissionResult.percent    ? String(commissionResult.percent)    : '',
    amountBYN:   commissionResult.amountBYN  || '',
    amountWords: commissionResult.amountWords || '',
    baseValue:   String(window.COMMISSION_CONFIG.baseValue),
    baseUnits:   commissionResult.baseUnits
                   ? commissionResult.baseUnits.toFixed(2).replace(/\.?0+$/, '')
                   : '',
  };

  // ── Задаток ─────────────────────────────────────────────────
  const depositBYNRaw = (getField('deal-Сумма задатка BYN') || '').replace(',', '.').trim();
  const depositUSDRaw = (getField('deal-Сумма задатка USD') || '').replace(',', '.').trim();
  const deposit = {
    amountBYN:      depositBYNRaw,
    amountBYNWords: depositBYNRaw ? window.moneyToText(depositBYNRaw)    : '',
    amountUSD:      depositUSDRaw,
    amountUSDWords: depositUSDRaw ? window.moneyToTextUSD(depositUSDRaw) : '',
  };

  // ── Остаточная стоимость после задатка (USD) ────────────────
  // remainderUSD = Стоимость USD − Сумма задатка USD
  // Вычисляется в памяти; в Excel не хранится.
  const _priceUSDNum   = parseFloat((getField('deal-Стоимость USD') || '0').replace(',', '.')) || 0;
  const _depositUSDNum = parseFloat(depositUSDRaw || '0') || 0;
  const _remainderUSD  = Math.max(0, _priceUSDNum - _depositUSDNum);
  property.remainderUSD      = _remainderUSD ? String(_remainderUSD) : '';
  property.remainderUSDWords = _remainderUSD ? window.moneyToTextUSD(_remainderUSD) : '';

  return { deal, property, seller, owner1, owner2, owner3, buyer, agent, agency, keys, money, commission, deposit };
}

// ============================================================
const TEMPLATE_REGISTRY = {

  'doverennost-pnd': {
    label: 'Доверенность ПНД',
    async generate(outputDir, options) {
      return window.electronAPI.generateDoverennost(buildPlaceholderData(), outputDir, options);
    },
  },

  'raspiska-klyuchi': {
    label: 'Расписка в получении ключей',
    async generate(outputDir, options) {
      return window.electronAPI.generateRaspiska(buildPlaceholderData(), outputDir, options);
    },
  },

  'reklama': {
    label: 'Договор на оказание рекламных услуг',
    async generate(outputDir, options) {
      return window.electronAPI.generateReklama(buildPlaceholderData(), outputDir, options);
    },
  },

  'rastorzhenie': {
    label: 'Соглашение о расторжении',
    async generate(outputDir, options) {
      return window.electronAPI.generateRastorzhenie(buildPlaceholderData(), outputDir, options);
    },
  },

  'zapros-pnd': {
    label: 'Запрос на ПНД',
    async generate(outputDir, options) {
      return window.electronAPI.generateZaprosPnd(buildPlaceholderData(), outputDir, options);
    },
  },

  'zapros-rsc': {
    label: 'Запрос в РСЦ',
    async generate(outputDir, options) {
      return window.electronAPI.generateZaprosRsc(buildPlaceholderData(), outputDir, options);
    },
  },

  'soglasie-obrabotka': {
    label: 'Согласие на обработку данных',
    async generate(outputDir, options) {
      return window.electronAPI.generateSoglasie(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-1-eksklyuziv': {
    label: 'Договор оказания риэлтерских услуг (1 собственник, эксклюзив)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp1Eksklyuziv(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-2-eksklyuziv': {
    label: 'Договор оказания риэлтерских услуг ЭКС (2 собственника, общий)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp2Eksklyuziv(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-2-obshiy': {
    label: 'Договор оказания риэлтерских услуг (2 собственника, общий)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp2Obshiy(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-3-eksklyuziv': {
    label: 'Договор оказания риэлтерских услуг (3 собственника, эксклюзив)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp3Eksklyuziv(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-3-obshiy': {
    label: 'Договор оказания риэлтерских услуг (3 собственника, общий)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp3Obshiy(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-1-obshiy': {
    label: 'Договор оказания риэлтерских услуг (1 собственник, общий)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkp1Obshiy(buildPlaceholderData(), outputDir, options);
    },
  },

  'konvertaciya': {
    label: 'Договор о конвертации валюты',
    async generate(outputDir, options) {
      return window.electronAPI.generateKonvertaciya(buildPlaceholderData(), outputDir, options);
    },
  },

  'zadatok-standart': {
    label: 'Договор задатка (стандартный)',
    async generate(outputDir, options) {
      return window.electronAPI.generateZadatokStandart(buildPlaceholderData(), outputDir, options);
    },
  },

  'dkp-fizlit-komstr': {
    label: 'Договор оказания риэлтерских услуг (физическое лицо — коммерческая структура)',
    async generate(outputDir, options) {
      return window.electronAPI.generateDkpFizlitKomstr(buildPlaceholderData(), outputDir, options);
    },
  },

};

// ============================================================
//  Generate documents
// ============================================================
btnGenerate.addEventListener('click', handleGenerate);

async function handleGenerate() {
  const outputDir = saveFolderInput.value.trim() || null;
  const options   = { addDate: !!(chkAddDate && chkAddDate.checked) };

  // Validate: save folder must be selected
  if (!outputDir) {
    showToast('✖ Сначала выберите папку для сохранения документов', 'error');
    saveFolderInput.classList.add('input-error-highlight');
    saveFolderInput.focus();
    setTimeout(() => saveFolderInput.classList.remove('input-error-highlight'), 2500);
    return;
  }

  // Collect checked, enabled checkboxes that have a data-template
  const checked = [...document.querySelectorAll(
    '.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]:checked'
  )];

  if (checked.length === 0) {
    showToast('✖ Выберите хотя бы один шаблон', 'error');
    return;
  }

  // Filter to only templates that are implemented in the registry
  const toGenerate = checked
    .map(cb => cb.closest('.tpl-item')?.dataset.template)
    .filter(key => key && TEMPLATE_REGISTRY[key]);

  if (toGenerate.length === 0) {
    showToast('Выбранные шаблоны ещё не реализованы');
    return;
  }

  let successCount = 0;
  const errors = [];

  for (const key of toGenerate) {
    const entry = TEMPLATE_REGISTRY[key];
    try {
      const result = await entry.generate(outputDir, options);
      if (result && result.success) {
        successCount++;
        if (chkOpenAfter && chkOpenAfter.checked) {
          window.electronAPI.openFile(result.path);
        }
      } else {
        errors.push(`${entry.label}: ${result?.error || 'неизвестная ошибка'}`);
      }
    } catch (err) {
      errors.push(`${entry.label}: ${err.message}`);
    }
  }

  if (successCount > 0) {
    showToast(`✔ Сформировано: ${successCount} из ${toGenerate.length}`);
  }
  errors.forEach(msg => showToast(`✖ ${msg}`, 'error'));
}

// ============================================================
//  Preview modal
// ============================================================
const previewOverlay       = document.getElementById('preview-overlay');
const previewTabs          = document.getElementById('preview-tabs');
const previewContent       = document.getElementById('preview-content');
const previewLoader        = document.getElementById('preview-loader');
const previewCloseBtn      = document.getElementById('preview-close');
const previewCloseFooter   = document.getElementById('btn-preview-close-footer');

function closePreviewModal() {
  previewOverlay.hidden = true;
  previewTabs.innerHTML = '';
  previewContent.innerHTML = '';
}

function renderParagraphs(paragraphs) {
  const page = document.createElement('div');
  page.className = 'preview-page';

  for (const text of paragraphs) {
    const p = document.createElement('div');
    p.className = 'preview-para' + (text.trim() === '' ? ' empty' : '');
    p.textContent = text || '';
    page.appendChild(p);
  }

  previewContent.innerHTML = '';
  previewContent.appendChild(page);
}

async function loadPreviewTab(templateKey, data) {
  previewContent.innerHTML = '';
  previewLoader.hidden = false;

  try {
    const result = await window.electronAPI.previewDocument(templateKey, data);
    previewLoader.hidden = true;

    if (!result || !result.success) {
      previewContent.innerHTML =
        `<div style="padding:32px;color:var(--error-text);font-size:13px;">
           ✖ Ошибка предпросмотра: ${result?.error || 'неизвестная ошибка'}
         </div>`;
      return;
    }

    renderParagraphs(result.paragraphs);
  } catch (err) {
    previewLoader.hidden = true;
    previewContent.innerHTML =
      `<div style="padding:32px;color:var(--error-text);font-size:13px;">✖ ${err.message}</div>`;
  }
}

function openPreviewModal(templateKeys) {
  const data = buildPlaceholderData();

  // Build tabs
  previewTabs.innerHTML = '';
  templateKeys.forEach((key, idx) => {
    const entry = TEMPLATE_REGISTRY[key];
    const btn = document.createElement('button');
    btn.className = 'preview-tab' + (idx === 0 ? ' active' : '');
    btn.textContent = entry.label;
    btn.addEventListener('click', () => {
      previewTabs.querySelectorAll('.preview-tab').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      loadPreviewTab(key, data);
    });
    previewTabs.appendChild(btn);
  });

  previewOverlay.hidden = false;
  previewContent.innerHTML = '';
  loadPreviewTab(templateKeys[0], data);
}

previewCloseBtn.addEventListener('click', closePreviewModal);
previewCloseFooter.addEventListener('click', closePreviewModal);
previewOverlay.addEventListener('click', (e) => {
  if (e.target === previewOverlay) closePreviewModal();
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !previewOverlay.hidden) closePreviewModal();
});

btnPreview.addEventListener('click', () => {
  const checked = [...document.querySelectorAll(
    '.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]:checked'
  )];

  const toPreview = checked
    .map(cb => cb.closest('.tpl-item')?.dataset.template)
    .filter(key => key && TEMPLATE_REGISTRY[key]);

  if (toPreview.length === 0) {
    showToast('✖ Выберите хотя бы один реализованный шаблон', 'error');
    return;
  }

  openPreviewModal(toPreview);
});

// ============================================================
//  О программе — модальное окно
// ============================================================
(function () {
  const overlay   = document.getElementById('about-overlay');
  const btnOpen   = document.getElementById('btn-about');
  const btnClose  = document.getElementById('about-close');
  const btnOk     = document.getElementById('about-ok');
  const siteLink  = document.getElementById('about-site-link');
  const emailLink = document.getElementById('about-email-link');

  function openAbout()  { overlay.hidden = false; }
  function closeAbout() { overlay.hidden = true; }

  btnOpen.addEventListener('click', openAbout);
  btnClose.addEventListener('click', closeAbout);
  btnOk.addEventListener('click', closeAbout);

  // Закрытие по клику на фон
  overlay.addEventListener('click', e => { if (e.target === overlay) closeAbout(); });

  // Закрытие по Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !overlay.hidden) closeAbout();
  });

  // Открытие ссылок через Electron shell
  if (window.electronAPI?.openExternal) {
    siteLink?.addEventListener('click', e => {
      e.preventDefault();
      window.electronAPI.openExternal('https://germesgarant.by');
    });
    emailLink?.addEventListener('click', e => {
      e.preventDefault();
      window.electronAPI.openExternal('mailto:mail@germesgarant.by');
    });
  }
}());

// ============================================================
//  Автообновление — модальное окно (Portable)
// ============================================================
(function () {
  if (!window.electronAPI?.onUpdateAvailable) return;

  const overlay      = document.getElementById('update-overlay');
  const stepNotify   = document.getElementById('update-step-notify');
  const stepProgress = document.getElementById('update-step-progress');
  const stepError    = document.getElementById('update-step-error');

  const versionLabel   = document.getElementById('update-version');
  const progressFill   = document.getElementById('update-progress-fill');
  const progressPct    = document.getElementById('update-progress-pct');
  const errorText      = document.getElementById('update-error-text');

  const btnConfirm    = document.getElementById('update-btn-confirm');
  const btnCancel     = document.getElementById('update-btn-cancel');
  const btnErrorClose = document.getElementById('update-btn-error-close');

  function showStep(step) {
    stepNotify.hidden   = (step !== 'notify');
    stepProgress.hidden = (step !== 'progress');
    stepError.hidden    = (step !== 'error');
  }

  function openModal() { overlay.hidden = false; }
  function closeModal() { overlay.hidden = true; }

  // Закрытие по клику на фон (только на шаге уведомления)
  overlay.addEventListener('click', e => {
    if (e.target === overlay && !stepNotify.hidden) closeModal();
  });

  // Закрытие по Escape (только на шаге уведомления)
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !overlay.hidden && !stepNotify.hidden) closeModal();
  });

  // ── Получено событие «Найдено обновление» ─────────────────
  window.electronAPI.onUpdateAvailable(({ version }) => {
    if (versionLabel) versionLabel.textContent = `v${version}`;
    showStep('notify');
    openModal();
  });

  // ── Пользователь нажал «Обновить» ─────────────────────────
  btnConfirm?.addEventListener('click', () => {
    showStep('progress');
    window.electronAPI.startUpdate();
  });

  // ── Пользователь нажал «Позже» ────────────────────────────
  btnCancel?.addEventListener('click', closeModal);

  // ── Прогресс скачивания ────────────────────────────────────
  window.electronAPI.onUpdateDownloadProgress(({ percent }) => {
    const pct = Math.min(100, Math.max(0, percent));
    if (progressFill) progressFill.style.width = `${pct}%`;
    if (progressPct)  progressPct.textContent   = `${pct}%`;
  });

  // ── Ошибка при обновлении ─────────────────────────────────
  window.electronAPI.onUpdateError(({ message }) => {
    if (errorText) errorText.textContent = message;
    showStep('error');
  });

  btnErrorClose?.addEventListener('click', closeModal);
}());
