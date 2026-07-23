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
let autoSaveTimer   = null;

// ============================================================
//  Persistent settings — папка и чекбоксы (localStorage)
// ============================================================
const LS_SETTINGS_KEY = 'germesSettings_v1';

function saveAppSettings() {
  try {
    localStorage.setItem(LS_SETTINGS_KEY, JSON.stringify({
      saveFolder:  saveFolderInput ? saveFolderInput.value : '',
      openAfter:   chkOpenAfter  ? chkOpenAfter.checked  : true,
      addDate:     chkAddDate    ? chkAddDate.checked     : false,
    }));
  } catch (_) {}
}

function restoreAppSettings() {
  try {
    const raw = localStorage.getItem(LS_SETTINGS_KEY);
    if (!raw) return;
    const s = JSON.parse(raw);
    if (saveFolderInput && typeof s.saveFolder === 'string' && s.saveFolder) {
      saveFolderInput.value = s.saveFolder;
    }
    if (chkOpenAfter  && typeof s.openAfter === 'boolean') chkOpenAfter.checked  = s.openAfter;
    if (chkAddDate    && typeof s.addDate   === 'boolean') chkAddDate.checked    = s.addDate;
  } catch (_) {}
}

// ============================================================
//  Universal helpers
// ============================================================
function isFieldEmpty(value) {
  if (value === null || value === undefined) return true;
  return String(value).trim() === '';
}

// ============================================================
//  Numeric input formatter — visual thousands separator (spaces)
//  Only reformats the display; all read/save helpers strip spaces
//  before processing so the stored value is always a plain number.
// ============================================================
function applyNumericFormat(input) {
  const selStart = input.selectionStart;
  const raw = input.value;
  // Remove all existing thousand-separating spaces
  const clean = raw.replace(/\s/g, '');

  if (clean === '') { input.value = ''; return; }
  // Allow only digits with an optional single decimal point
  if (!/^\d*\.?\d*$/.test(clean)) return;

  const dotIdx = clean.indexOf('.');
  const intPart = dotIdx >= 0 ? clean.slice(0, dotIdx) : clean;
  const decPart = dotIdx >= 0 ? clean.slice(dotIdx)    : '';

  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  const newVal = formattedInt + decPart;

  if (newVal === raw) return;
  input.value = newVal;

  // Restore cursor: count non-space chars before old cursor, find same count in new value
  const charsBeforeCursor = raw.slice(0, selStart).replace(/\s/g, '').length;
  let newCursor = newVal.length;
  let counted = 0;
  for (let i = 0; i < newVal.length; i++) {
    if (newVal[i] !== ' ') counted++;
    if (counted === charsBeforeCursor) { newCursor = i + 1; break; }
  }
  try { input.setSelectionRange(newCursor, newCursor); } catch (_) {}
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
  // Apply thousands formatting immediately on load for numeric fields
  if (el.dataset.numeric && el.value) applyNumericFormat(el);
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
  // Enable Save As when a file is loaded OR when there are filled fields without a file
  if (hasDirty && !currentFilePath) btnSaveAs.disabled = false;
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
  // Автосохранение: запускаем/сбрасываем дебаунс только если файл уже открыт
  if (currentFilePath) {
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(autoSave, 30000);
  }
  // Re-evaluate contract availability whenever owner fields or seller-ownership flag changes
  if (
    inputId.startsWith('owner1-') ||
    inputId.startsWith('owner2-') ||
    inputId.startsWith('owner3-') ||
    inputId === 'seller-Является собственником'
  ) {
    updateContractAvailability();
  }
  // Re-evaluate object-type-dependent field visibility
  if (inputId === 'property-Тип объекта') {
    applyObjectTypeVisibility();
    autoUpdateCommission();
  }
  // Update block completion badge + object preview
  const dashIdx2 = inputId.indexOf('-');
  if (dashIdx2 !== -1) {
    const pfx = inputId.slice(0, dashIdx2);
    updateBlockCompletion(pfx);
    if (pfx === 'property' || pfx === 'deal') updateObjPreview();
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
  if (e.target.dataset.numeric) applyNumericFormat(e.target);
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
    if (el) updates[rowNum] = el.dataset.numeric
      ? el.value.replace(/\s/g, '').trim()
      : el.value.trim();
  }
  return updates;
}

// ============================================================
//  Build field groups for creating Excel from scratch
//  Returns { deal: {fieldKey: value}, property: {...}, ... }
// ============================================================
function buildFieldGroups() {
  const groups = { deal: {}, property: {}, seller: {}, owner1: {}, owner2: {}, owner3: {}, buyer: {} };
  for (const [mapKey, inputId] of Object.entries(FIELD_MAP)) {
    const dashIdx = mapKey.indexOf('-');
    if (dashIdx === -1) continue;
    const blockId  = mapKey.slice(0, dashIdx);
    const fieldKey = mapKey.slice(dashIdx + 1);
    if (!groups[blockId]) continue;
    const el = document.getElementById(inputId);
    if (el) groups[blockId][fieldKey] = el.dataset.numeric
      ? el.value.replace(/\s/g, '').trim()
      : el.value.trim();
  }
  return groups;
}

// ============================================================
//  Default "Save As" filename
// ============================================================
function buildDefaultSaveAsName() {
  const contractNum = (document.getElementById('deal-Номер договора')?.value || '').trim();
  const address     = (document.getElementById('property-Адрес')?.value || '').trim();
  const dealDate    = (document.getElementById('deal-Дата договора')?.value || '').trim();

  // Sanitize a string for use in a filename: replace characters forbidden on Windows/macOS
  function sanitize(str) {
    return str.replace(/[/\\:*?"<>|]/g, '-').replace(/-+/g, '-').trim();
  }

  const parts = [];

  if (contractNum || address) {
    if (contractNum) parts.push(sanitize(contractNum));
    if (address)     parts.push(sanitize(address));
  } else if (dealDate) {
    // Fallback: convert dd.mm.yyyy → yyyy-mm-dd for a clean filename
    const dateParts = dealDate.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    parts.push(dateParts ? `${dateParts[3]}-${dateParts[2]}-${dateParts[1]}` : sanitize(dealDate));
  } else {
    parts.push('Сделка');
  }

  return parts.join('_') + '.xlsx';
}

// ============================================================
//  Save
// ============================================================
async function handleSave() {
  if (dirtyInputIds.size === 0) return;
  // No file loaded — delegate to Save As which will create a new file from scratch
  if (!currentFilePath) {
    await handleSaveAs();
    return;
  }
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
//  Auto-save (debounced, only when file is already open)
// ============================================================
async function autoSave() {
  if (!currentFilePath || dirtyInputIds.size === 0) return;
  setStatus('Автосохранение…');
  try {
    const updates = buildUpdates();
    await window.electronAPI.writeExcel(currentFilePath, currentFilePath, updates);
    commitCurrentValues();
    const now = new Date();
    const hhmm = now.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    setStatus(`Автосохранено в ${hhmm}`);
  } catch (err) {
    setStatus('Ошибка автосохранения: ' + err.message);
  }
}

// ============================================================
//  Save As
// ============================================================
async function handleSaveAs() {
  const defaultName = buildDefaultSaveAsName();
  const defaultPath = currentFilePath
    ? currentFilePath.replace(/[^/\\]+$/, defaultName)
    : defaultName;

  let targetPath;
  try {
    targetPath = await window.electronAPI.saveFileDialog(defaultPath);
  } catch (err) {
    showToast('✖ Ошибка при открытии диалога: ' + err.message, 'error');
    return;
  }
  if (!targetPath) return;

  try {
    if (currentFilePath) {
      // Existing file loaded — copy its structure and write updates
      const updates = buildUpdates();
      await window.electronAPI.writeExcel(currentFilePath, targetPath, updates);
    } else {
      // No file loaded — build Excel from scratch using fields-config structure
      const fieldGroups = buildFieldGroups();
      const result = await window.electronAPI.createExcelFromData(fieldGroups, targetPath);
      if (!result.ok) throw new Error('Не удалось создать файл');
      // Bind the new rowMap so subsequent saves work via writeExcel normally
      rowMap = result.rowMap;
    }

    currentFilePath = targetPath;
    filePathDisplay.value = targetPath;
    const baseName = targetPath.split(/[\\/]/).pop();
    dropFileName.textContent = baseName;
    setDropState('success');
    btnSaveAs.disabled = false;
    commitCurrentValues();
    setStatus('Файл сохранён: ' + baseName);
    showToast('✔ Файл успешно сохранён');

    // История последних документов
    window.RecentDocs?.push({
      name:  baseName,
      type:  'excel',
      icon:  'excel',
      path:  targetPath,
      label: baseName,
    });
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
  applyObjectTypeVisibility();
  // Уведомить UIController об обновлении формы
  document.dispatchEvent(new Event('form:populated'));
  updateBlockCompletion(null); // пересчитать все блоки после загрузки
  updateObjPreview();
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
  applyObjectTypeVisibility();
  // Уведомить UIController об очистке формы
  document.dispatchEvent(new Event('form:cleared'));
  updateBlockCompletion(null); // сбросить все бейджи после очистки
  updateObjPreview();
}

// ============================================================
//  Check data (basic validation)
// ============================================================

// Вспомогательная: полный набор полей собственника N
function ownerCheckFields(prefix, n) {
  return [
    [`${prefix}-Фамилия`,                 `Собственник ${n} → Фамилия`],
    [`${prefix}-Имя`,                     `Собственник ${n} → Имя`],
    [`${prefix}-Отчество`,                `Собственник ${n} → Отчество`],
    [`${prefix}-Дата рождения`,           `Собственник ${n} → Дата рождения`],
    [`${prefix}-Паспорт серия`,           `Собственник ${n} → Паспорт серия`],
    [`${prefix}-Паспорт номер`,           `Собственник ${n} → Паспорт номер`],
    [`${prefix}-Идентификационный номер`, `Собственник ${n} → Идент. номер`],
    [`${prefix}-Кем выдан`,              `Собственник ${n} → Кем выдан`],
    [`${prefix}-Дата выдачи`,            `Собственник ${n} → Дата выдачи`],
    [`${prefix}-Адрес регистрации`,       `Собственник ${n} → Адрес регистрации`],
    [`${prefix}-Доля собственности`,      `Собственник ${n} → Доля собственности`],
  ];
}

function handleCheckData() {
  // ── Базовые проверки из Smart Panel (getIssues — ui-controller.js) ─
  // getIssues() доступна глобально: ui-controller.js загружается сразу
  // после app.js, и к моменту клика оба скрипта уже выполнены.
  const smartIssues = getIssues();
  const missing     = smartIssues.map(f => f.label);
  const checkedIds  = new Set(smartIssues.map(f => f.id));

  // chk добавляет поле только если его ещё нет в getIssues
  function chk(id, label) {
    if (checkedIds.has(id)) return;
    const el = document.getElementById(id);
    if (el && isFieldEmpty(el.value)) missing.push(label);
  }

  // ── Дополнительные поля: только кнопка, нет в Smart Panel ────
  chk('deal-Количество собственников', 'Сделка → Кол-во собственников');
  chk('deal-Ответственный риэлтер',    'Сделка → Риэлтер');

  // Тип объекта нестандартный / не выбран → проверяем оба идентификатора
  // (getIssues не имеет else-ветки для неизвестного типа)
  const propTypeRaw = (getField('property-Тип объекта') || '').trim().toLowerCase();
  const isHouseChk  = propTypeRaw === 'дом' || propTypeRaw === 'жилой дом';
  const isFlatChk   = propTypeRaw === 'квартира' || propTypeRaw === 'апартаменты' || propTypeRaw === 'комната';
  const isCommerChk = propTypeRaw === 'коммерческая недвижимость';
  if (!isHouseChk && !isFlatChk && !isCommerChk) {
    chk('property-Кадастровый номер', 'Объект → Кадастровый №');
    chk('property-Инвентарный номер', 'Объект → Инвентарный №');
  }

  // Покупатель — при ручной проверке всегда обязателен
  // (Smart Panel требует покупателя только при наличии задатка)
  const depositBYN = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
  const depositUSD = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
  if (depositBYN === '' && depositUSD === '') {
    [
      ['buyer-Фамилия',                 'Покупатель → Фамилия'],
      ['buyer-Имя',                     'Покупатель → Имя'],
      ['buyer-Отчество',                'Покупатель → Отчество'],
      ['buyer-Дата рождения',           'Покупатель → Дата рождения'],
      ['buyer-Паспорт серия',           'Покупатель → Паспорт серия'],
      ['buyer-Паспорт номер',           'Покупатель → Паспорт номер'],
      ['buyer-Идентификационный номер', 'Покупатель → Идент. номер'],
      ['buyer-Кем выдан',              'Покупатель → Кем выдан'],
      ['buyer-Дата выдачи',            'Покупатель → Дата выдачи'],
      ['buyer-Адрес регистрации',       'Покупатель → Адрес регистрации'],
    ].forEach(([id, label]) => chk(id, label));
  }

  // Собственник 1 — расширенный набор
  // (Smart Panel проверяет только 6 базовых полей из OWNER1_FIELDS)
  const isOwnerRaw = (getField('seller-Является собственником') || '').trim().toLowerCase();
  if (isOwnerPresent('owner1') || isOwnerRaw === 'нет') {
    [
      ['owner1-Отчество',           'Собственник 1 → Отчество'],
      ['owner1-Дата рождения',      'Собственник 1 → Дата рождения'],
      ['owner1-Кем выдан',         'Собственник 1 → Кем выдан'],
      ['owner1-Дата выдачи',       'Собственник 1 → Дата выдачи'],
      ['owner1-Доля собственности', 'Собственник 1 → Доля собственности'],
    ].forEach(([id, label]) => chk(id, label));
  }

  // Собственники 2 и 3 — полный набор (Smart Panel их не проверяет)
  if (isOwnerPresent('owner2')) ownerCheckFields('owner2', 2).forEach(([id, label]) => chk(id, label));
  if (isOwnerPresent('owner3')) ownerCheckFields('owner3', 3).forEach(([id, label]) => chk(id, label));

  // ── Результат ─────────────────────────────────────────────────
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

  // Normalise: strip thousand-separator spaces, treat comma as decimal separator
  const raw = bynInput.value.replace(/\s/g, '').replace(',', '.').trim();

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

  const raw = (bynInput ? bynInput.value : '').replace(/\s/g, '').replace(',', '.').trim();
  if (raw === '' || !/^\d+(\.\d{1,2})?$/.test(raw)) {
    commissionInput.value = '';
    return;
  }

  const propType = (document.getElementById('property-Тип объекта')?.value || '').trim().toLowerCase();
  const isCommercial = propType === 'коммерческая недвижимость';
  const cfg = isCommercial ? window.COMMISSION_CONFIG_COMMERCIAL : window.COMMISSION_CONFIG;

  const result = window.calculateCommission(parseFloat(raw), cfg.baseValue, cfg.brackets);
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
//  Object preview bar (above accordion)
// ============================================================
const OBJ_ICON_SVG = {
  'дом':      '<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15 16 5l12 10"/><path d="M7 13v14h18V13"/><rect x="13" y="19" width="6" height="8" rx="1"/></svg>',
  'жилой дом':'<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15 16 5l12 10"/><path d="M7 13v14h18V13"/><rect x="13" y="19" width="6" height="8" rx="1"/></svg>',
  'квартира': '<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="24" height="24" rx="2"/><line x1="4" y1="13" x2="28" y2="13"/><line x1="4" y1="22" x2="28" y2="22"/><line x1="13" y1="4" x2="13" y2="28"/><line x1="22" y1="4" x2="22" y2="28"/></svg>',
  'апартаменты':'<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="4" width="24" height="24" rx="2"/><line x1="4" y1="13" x2="28" y2="13"/><line x1="4" y1="22" x2="28" y2="22"/><line x1="13" y1="4" x2="13" y2="28"/><line x1="22" y1="4" x2="22" y2="28"/></svg>',
  'комната':  '<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="7" y="3" width="18" height="26" rx="2"/><circle cx="21" cy="16" r="1.5" fill="currentColor" stroke="none"/><line x1="7" y1="29" x2="25" y2="29"/></svg>',
  'коммерческая недвижимость':'<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="10" width="26" height="19" rx="1"/><path d="M3 10 9 4h14l6 6"/><rect x="9" y="17" width="5" height="12"/><rect x="18" y="17" width="5" height="8"/></svg>',
  'default':  '<svg width="26" height="26" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="11" width="24" height="18" rx="2"/><path d="M4 13 16 4l12 9"/><line x1="13" y1="29" x2="13" y2="20"/><line x1="19" y1="29" x2="19" y2="20"/><line x1="13" y1="20" x2="19" y2="20"/></svg>',
};

/** Форматирует число с пробелами как тысячный разделитель (123 456.78) */
function fmtPrice(raw) {
  const num = parseFloat(String(raw).replace(/\s/g, '').replace(',', '.'));
  if (isNaN(num)) return raw;
  const [int, dec] = num.toFixed(num % 1 !== 0 ? 2 : 0).split('.');
  const intFmt = int.replace(/\B(?=(\d{3})+(?!\d))/g, '\u00A0'); // non-breaking space
  return dec ? `${intFmt}.${dec}` : intFmt;
}

function updateObjPreview() {
  const bar = document.getElementById('obj-preview-bar');
  if (!bar) return;

  const type   = (getField('property-Тип объекта')     || '').trim();
  const city   = (getField('property-Город')           || '').trim();
  const street = (getField('property-Улица')           || '').trim();
  const house  = (getField('property-Дом')             || '').trim();
  const corps  = (getField('property-Корпус')          || '').trim();
  const apt    = (getField('property-Квартира')        || '').trim();
  const rooms  = (getField('property-Количество комнат')|| '').trim();
  const area   = (getField('property-Общая площадь')   || '').trim();
  const rawBYN = (getField('deal-Стоимость BYN')       || '').replace(/\s/g, '').trim();
  const rawUSD = (getField('deal-Стоимость USD')       || '').replace(/\s/g, '').trim();

  const hasData = type || city || street || house || rawBYN || rawUSD;
  bar.classList.toggle('obj-preview-bar--empty', !hasData);

  // Trigger subtle blink animation on update
  bar.classList.remove('obj-preview-bar--updating');
  void bar.offsetWidth;
  if (hasData) bar.classList.add('obj-preview-bar--updating');

  // Icon
  const iconEl = document.getElementById('obj-preview-icon');
  if (iconEl) {
    const key = type.toLowerCase();
    iconEl.innerHTML = OBJ_ICON_SVG[key] || OBJ_ICON_SVG['default'];
  }

  // Title
  const titleEl = document.getElementById('obj-preview-title');
  if (titleEl) titleEl.textContent = type ? `Объект: ${type}` : 'Объект не выбран';

  // Address
  const addrEl = document.getElementById('obj-preview-address');
  if (addrEl) {
    const parts = [];
    if (city) parts.push(city);
    if (street) {
      let s = `ул. ${street}`;
      if (house) {
        s += `, ${house}`;
        if (corps) s += ` корп.\u00A0${corps}`;
        if (apt)   s += `-${apt}`;
      }
      parts.push(s);
    }
    addrEl.textContent = parts.join(', ');
  }

  // Details chips
  const detailEl = document.getElementById('obj-preview-details');
  if (detailEl) {
    const chips = [];
    if (area)  chips.push(`${area}\u00A0м²`);
    if (rooms) chips.push(`${rooms}\u00A0комн.`);
    detailEl.textContent = chips.join('  ·  ');
  }

  // Price BYN
  const bynEl = document.getElementById('obj-preview-byn');
  if (bynEl) bynEl.textContent = rawBYN ? `${fmtPrice(rawBYN)}\u00A0BYN` : '—';

  // Price USD
  const usdEl = document.getElementById('obj-preview-usd');
  if (usdEl) usdEl.textContent = rawUSD ? `≈\u00A0${fmtPrice(rawUSD)}\u00A0USD` : '';
}

// ============================================================
//  Block completion badges + glow animation
// ============================================================
const BLOCK_COMPLETE_LABELS = {
  deal:     'Сделка заполнена',
  property: 'Объект заполнен',
  seller:   'Продавец заполнен',
  owners:   'Собственники заполнены',
  buyer:    'Покупатель заполнен',
};

/** true если el виден в layout (нет скрытого предка) */
function isInputVisible(el) {
  // Inputs в скрытых tab-pane (display:none) имеют offsetParent === null
  if (el.offsetParent === null && getComputedStyle(el).position !== 'fixed') return false;
  // Поля, скрытые по типу объекта — display:none на .fr-строке
  const row = el.closest('.fr');
  if (row && row.style.display === 'none') return false;
  return true;
}

/** true если все user-editable видимые поля блока заполнены */
function isBlockFilled(prefix) {
  const ids = Object.entries(FIELD_MAP)
    .filter(([k]) => k.startsWith(prefix + '-'))
    .map(([, id]) => id);
  if (!ids.length) return false;
  return ids.every(id => {
    const el = document.getElementById(id);
    if (!el)              return true; // нет в DOM — пропуск
    if (el.readOnly)      return true; // computed/readonly — пропуск
    if (!isInputVisible(el)) return true; // скрыт фильтром — пропуск
    return !isFieldEmpty(el.value);
  });
}

/** Обновляет badge и класс ws-block--complete для нужных блоков */
function updateBlockCompletion(changedPrefix) {
  const toCheck = new Set();
  if (!changedPrefix) {
    ['deal', 'property', 'seller', 'owners', 'buyer'].forEach(b => toCheck.add(b));
  } else if (changedPrefix.startsWith('owner')) {
    toCheck.add('owners');
  } else {
    toCheck.add(changedPrefix);
  }

  for (const blockId of toCheck) {
    const section = document.getElementById(`ws-${blockId}`);
    const badge   = document.getElementById(`blk-status-${blockId}`);
    if (!section || !badge) continue;

    let complete;
    if (blockId === 'owners') {
      const prefixes = ['owner1'];
      if (isOwnerPresent('owner2')) prefixes.push('owner2');
      if (isOwnerPresent('owner3')) prefixes.push('owner3');
      complete = prefixes.every(p => isBlockFilled(p));
    } else {
      complete = isBlockFilled(blockId);
    }

    const wasComplete = section.classList.contains('ws-block--complete');

    if (complete) {
      section.classList.add('ws-block--complete');
      badge.textContent = BLOCK_COMPLETE_LABELS[blockId] || '';
      if (!wasComplete) {
        // Перезапуск flash-анимации при переходе незаполнен → заполнен
        section.classList.remove('ws-block--complete-flash');
        void section.offsetWidth; // reflow
        section.classList.add('ws-block--complete-flash');
      }
    } else {
      section.classList.remove('ws-block--complete', 'ws-block--complete-flash');
      badge.textContent = '';
    }
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
//  Object-type-dependent field visibility
//  Поля с data-object-type="дом" — только для домов.
//  Поля с data-object-type="квартира" — только для квартир.
//  Когда тип не выбран — скрываем все условные поля.
// ============================================================
function applyObjectTypeVisibility() {
  const raw  = (getField('property-Тип объекта') || '').trim().toLowerCase();
  const isHouse      = raw === 'дом' || raw === 'жилой дом';
  const isFlat       = raw === 'квартира' || raw === 'апартаменты' || raw === 'комната';
  const isCommercial = raw === 'коммерческая недвижимость';
  const isEmpty = raw === '';

  document.querySelectorAll('[data-object-type]').forEach((el) => {
    const type = el.dataset.objectType;
    if (isEmpty) {
      el.style.display = 'none';
      return;
    }
    if (type === 'дом') {
      el.style.display = isHouse ? '' : 'none';
    } else if (type === 'квартира') {
      el.style.display = isFlat ? '' : 'none';
    } else if (type === 'коммерческая недвижимость') {
      el.style.display = isCommercial ? '' : 'none';
    }
  });
}

// ============================================================
//  Template checkboxes
// ============================================================
function handleSelectAll() {
  document.querySelectorAll('.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]').forEach((cb) => { cb.checked = true; });
  updateSidebarStatus();
}
function handleDeselectAll() {
  document.querySelectorAll('.tpl-item input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
  updateSidebarStatus();
}

// ============================================================
//  Sidebar status — count badge + warnings
// ============================================================
function updateSidebarStatus() {
  const warningsEl = document.getElementById('sidebar-warnings');
  const badgeEl    = document.getElementById('tpl-count-badge');
  if (!warningsEl || !badgeEl) return;

  const hasFolder = saveFolderInput.value.trim() !== '';
  const checkedCount = document.querySelectorAll(
    '.tpl-item:not(.tpl-item-disabled) input[type="checkbox"]:checked'
  ).length;

  // Count badge in the Templates header
  if (checkedCount > 0) {
    badgeEl.textContent = checkedCount;
    badgeEl.hidden = false;
  } else {
    badgeEl.hidden = true;
  }
}

// React to individual template checkbox changes
document.addEventListener('change', (e) => {
  if (e.target.matches('.tpl-item input[type="checkbox"]')) updateSidebarStatus();
});

// React to manual edits / clear of the folder input
saveFolderInput.addEventListener('input', () => { updateSidebarStatus(); saveAppSettings(); });

// Persist settings checkboxes on change
if (chkOpenAfter) chkOpenAfter.addEventListener('change', saveAppSettings);
if (chkAddDate)   chkAddDate.addEventListener('change', saveAppSettings);

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

  // История последних документов
  window.RecentDocs?.push({
    name:  baseName,
    type:  'excel',
    icon:  'excel',
    path:  filePath,
    label: baseName,
  });
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
if (btnCheck) btnCheck.addEventListener('click', handleCheckData);
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
  if (chosen) { saveFolderInput.value = chosen; updateSidebarStatus(); saveAppSettings(); }
});

// ============================================================
//  Template data helpers
// ============================================================
function getField(id) {
  return (document.getElementById(id)?.value || '').trim();
}

// Числовые поля: возвращает значение без пробелов-разделителей тысяч (для Word-документов)
function getNumericField(id) {
  return getField(id).replace(/\s/g, '');
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
  const isHouse      = propertyTypeRaw === 'дом';
  const isApartment  = propertyTypeRaw === 'квартира';
  const isCommercial = propertyTypeRaw === 'коммерческая недвижимость';

  const _endDateRaw = getField('deal-Дата окончания договора') || '';
  const deal = {
    number:                    getField('deal-Номер договора')  || '',
    date:                      getField('deal-Дата договора')   || '',
    dateText:                  '',
    endDate:                   _endDateRaw,
    endDateText:               dateToLongRussian(_endDateRaw),
    contractNumber:            getField('deal-Номер договора')  || '',
    contractDate:              getField('deal-Дата договора')   || '',
    advertisingContractNumber: '',
    advertisingContractDate:   '',
    depositContractNumber:     '',
    depositContractDate:       getField('deal-Дата договора задатка')          || '',
    depositContractEndDate:    getField('deal-Дата окончания договора задатка') || '',
    storageContractNumber:     '',
    storageContractDate:       '',
    additionalTerms:           getField('deal-Дополнительные условия') || '',
    btiPayment:                getField('deal-Оплата услуг БТИ') || '',
    furniture:                 getField('deal-Мебель')            || '',
  };

  const property = {
    isHouse,
    isApartment,
    isCommercial,
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
    landArea:        getField('property-Площадь участка')   || '',
    ownershipForm:   getField('property-Форма собственности') || '',
    inventoryNumber: getField('property-Инвентарный номер')|| '',
    wallMaterial:    getField('property-Материал стен')    || '',
    yearBuilt:    getField('property-Год постройки')   || '',
    commercialKind:    getField('property-Вид коммерческой недвижимости')       || '',
    commercialPurpose: getField('property-Назначение коммерческой недвижимости') || '',
    priceUSD:          getNumericField('deal-Стоимость USD'),
    priceBYN:          getNumericField('deal-Стоимость BYN'),
    priceWords:        getField('deal-Стоимость прописью') || '',
    priceWordsUSD:     (() => {
      const raw = getNumericField('deal-Стоимость USD').replace(',', '.');
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

  // ── Риэлтер — единый источник правды: RealtorService ──────
  // RealtorService является основным источником текущего риэлтера.
  // Поле «Ответственный риэлтер» из Excel используется только как
  // запасной вариант, если RealtorService недоступен.
  const _realtorRecord = window.RealtorService?.getCurrent?.() ?? null;
  const agentRecord = _realtorRecord || (() => {
    const agentRaw        = getField('deal-Ответственный риэлтер') || '';
    const agentNormalized = agentRaw.trim().toLowerCase();
    return (window.AGENTS_CONFIG?.agents || []).find((a) =>
      a.matchKeys.some((key) => agentNormalized.includes(key))
    ) || null;
  })();
  const _agentFallbackRaw = getField('deal-Ответственный риэлтер') || '';
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
        fullName:          _agentFallbackRaw,
        initials:          _agentFallbackRaw,
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
  // Для коммерческой недвижимости — отдельная тарифная таблица.
  // Для остальных объектов — стандартная.
  const priceBYNRaw = parseFloat(
    (getField('deal-Стоимость BYN') || '0').replace(',', '.')
  );
  const commissionCfg = isCommercial
    ? window.COMMISSION_CONFIG_COMMERCIAL
    : window.COMMISSION_CONFIG;
  const commissionResult = window.calculateCommission(
    priceBYNRaw,
    commissionCfg.baseValue,
    commissionCfg.brackets
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
  const depositBYNRaw = getNumericField('deal-Сумма задатка BYN').replace(',', '.');
  const deposit = {
    amountBYN:      depositBYNRaw,
    amountBYNWords: depositBYNRaw ? window.moneyToText(depositBYNRaw) : '',
    amountUSD:      '',
    amountUSDWords: '',
  };

  // remainderUSD не используется (задаток в USD отключён)
  property.remainderUSD      = '';
  property.remainderUSDWords = '';

  return { deal, property, seller, owner1, owner2, owner3, buyer, agent, agency, keys, money, commission, deposit };
}

// ============================================================
// Вспомогательная функция: создаёт метод generate для заданного ключа шаблона
function makeGenerate(key) {
  return async function(outputDir, options) {
    return window.electronAPI.generateDocument(key, buildPlaceholderData(), outputDir, options);
  };
}

const TEMPLATE_REGISTRY = {
  'doverennost-pnd':    { label: 'Доверенность ПНД',                                                                    generate: makeGenerate('doverennost-pnd') },
  'raspiska-klyuchi':   { label: 'Расписка в получении ключей',                                                          generate: makeGenerate('raspiska-klyuchi') },
  'reklama':            { label: 'Договор на оказание рекламных услуг',                                                  generate: makeGenerate('reklama') },
  'rastorzhenie':       { label: 'Соглашение о расторжении',                                                             generate: makeGenerate('rastorzhenie') },
  'zapros-pnd':         { label: 'Запрос на ПНД',                                                                        generate: makeGenerate('zapros-pnd') },
  'zapros-rsc':         { label: 'Запрос в РСЦ',                                                                         generate: makeGenerate('zapros-rsc') },
  'soglasie-obrabotka': { label: 'Согласие на обработку данных',                                                         generate: makeGenerate('soglasie-obrabotka') },
  'dkp-1-eksklyuziv':  { label: 'Договор оказания риэлтерских услуг (1 собственник, эксклюзив)',                         generate: makeGenerate('dkp-1-eksklyuziv') },
  'dkp-2-eksklyuziv':  { label: 'Договор оказания риэлтерских услуг ЭКС (2 собственника, общий)',                        generate: makeGenerate('dkp-2-eksklyuziv') },
  'dkp-2-obshiy':      { label: 'Договор оказания риэлтерских услуг (2 собственника, общий)',                            generate: makeGenerate('dkp-2-obshiy') },
  'dkp-3-eksklyuziv':  { label: 'Договор оказания риэлтерских услуг (3 собственника, эксклюзив)',                        generate: makeGenerate('dkp-3-eksklyuziv') },
  'dkp-3-obshiy':      { label: 'Договор оказания риэлтерских услуг (3 собственника, общий)',                            generate: makeGenerate('dkp-3-obshiy') },
  'dkp-1-obshiy':      { label: 'Договор оказания риэлтерских услуг (1 собственник, общий)',                             generate: makeGenerate('dkp-1-obshiy') },
  'konvertaciya':      { label: 'Договор о конвертации валюты',                                                          generate: makeGenerate('konvertaciya') },
  'zadatok-standart':  { label: 'Договор задатка (стандартный)',                                                         generate: makeGenerate('zadatok-standart') },
  'dkp-fizlit-komstr': { label: 'Договор оказания риэлтерских услуг (физическое лицо — коммерческая структура)',         generate: makeGenerate('dkp-fizlit-komstr') },
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
        // История последних документов
        window.RecentDocs?.push({
          name:  result.path.split(/[/\\]/).pop(),
          type:  'word',
          icon:  key,
          path:  result.path,
          label: entry.label,
        });
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

    const page = document.createElement('div');
    page.className = 'preview-page';
    page.innerHTML = result.html;
    previewContent.innerHTML = '';
    previewContent.appendChild(page);
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

  // Restore persistent settings (папка, чекбоксы)
  restoreAppSettings();
  // Initial sidebar state on app load
  updateSidebarStatus();
  // Скрываем условные поля при старте (тип объекта не выбран)
  applyObjectTypeVisibility();
}());
