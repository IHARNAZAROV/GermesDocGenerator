'use strict';

// ============================================================
//  Утилита: экранирование HTML-спецсимволов
// ============================================================
function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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
//  Modal controller — common Escape/overlay/focus handling
// ============================================================
class ModalController {
  constructor(overlay, options = {}) {
    this.overlay = overlay;
    this.dialog = options.dialog || overlay?.querySelector('[role="dialog"]') || overlay?.firstElementChild || overlay;
    this.closeOnOverlay = options.closeOnOverlay !== false;
    this.closeOnEscape = options.closeOnEscape !== false;
    this.initialFocus = options.initialFocus || null;
    this.onClose = options.onClose || null;
    this.shouldClose = options.shouldClose || null;
    this.isDynamic = !!options.isDynamic;
    this.isOpen = false;
    this.previousFocus = null;

    this._handleOverlayClick = this._handleOverlayClick.bind(this);
    this._handleKeydown = this._handleKeydown.bind(this);
  }

  open({ initialFocus } = {}) {
    if (!this.overlay || this.isOpen) return;
    this.isOpen = true;
    this.previousFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    this.overlay.hidden = false;
    this.overlay.addEventListener('click', this._handleOverlayClick);
    document.addEventListener('keydown', this._handleKeydown);
    requestAnimationFrame(() => this._focusInitial(initialFocus));
  }

  close(detail) {
    if (!this.overlay || !this.isOpen) return;
    if (this.shouldClose && !this.shouldClose(detail)) return;
    this.isOpen = false;
    this.overlay.removeEventListener('click', this._handleOverlayClick);
    document.removeEventListener('keydown', this._handleKeydown);
    if (this.isDynamic) {
      this.overlay.remove();
    } else {
      this.overlay.hidden = true;
    }
    this.onClose?.(detail);
    this._restoreFocus();
  }

  _handleOverlayClick(e) {
    if (this.closeOnOverlay && e.target === this.overlay) this.close({ reason: 'overlay' });
  }

  _handleKeydown(e) {
    if (e.key === 'Escape' && this.closeOnEscape) {
      e.stopPropagation();
      this.close({ reason: 'escape' });
      return;
    }
    if (e.key === 'Tab') this._trapFocus(e);
  }

  _focusInitial(initialFocus) {
    const target = this._resolveFocusTarget(initialFocus || this.initialFocus)
      || this._focusableElements()[0]
      || this.dialog;
    target?.focus?.();
  }

  _resolveFocusTarget(target) {
    if (!target) return null;
    if (typeof target === 'string') return this.overlay.querySelector(target);
    if (target instanceof HTMLElement) return target;
    return null;
  }

  _focusableElements() {
    if (!this.overlay) return [];
    return [...this.overlay.querySelectorAll('a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])')]
      .filter(el => !el.hidden && el.getClientRects().length > 0);
  }

  _trapFocus(e) {
    const focusable = this._focusableElements();
    if (focusable.length === 0) {
      e.preventDefault();
      this.dialog?.focus?.();
      return;
    }
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  }

  _restoreFocus() {
    if (this.previousFocus && document.contains(this.previousFocus)) {
      this.previousFocus.focus?.();
    }
    this.previousFocus = null;
  }
}

window.ModalController = ModalController;

// ============================================================
//  Application state
// ============================================================
let currentFilePath = null;
let currentFileToken = null;
let saveFolderToken = null;
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
    // Собираем состояние всех чекбоксов шаблонов
    const tplChecked = {};
    document.querySelectorAll('.tpl-item[data-template] input[type="checkbox"]').forEach((cb) => {
      const key = cb.closest('.tpl-item')?.dataset.template;
      if (key) tplChecked[key] = cb.checked;
    });

    localStorage.setItem(LS_SETTINGS_KEY, JSON.stringify({
      saveFolder:  saveFolderInput ? saveFolderInput.value : '',
      openAfter:   chkOpenAfter  ? chkOpenAfter.checked  : true,
      addDate:     chkAddDate    ? chkAddDate.checked     : false,
      tplChecked,
    }));
  } catch (_) {}
}

function restoreAppSettings() {
  try {
    const raw = localStorage.getItem(LS_SETTINGS_KEY);
    if (!raw) return;
    const s = JSON.parse(raw);

    // Restore folder display text immediately so UI shows the path
    if (saveFolderInput && typeof s.saveFolder === 'string' && s.saveFolder) {
      saveFolderInput.value = s.saveFolder;

      // Re-mint a valid session token for the saved path (no dialog needed)
      if (window.electronAPI?.restoreFolder) {
        window.electronAPI.restoreFolder(s.saveFolder).then((result) => {
          if (result) {
            saveFolderToken = result.token;
          } else {
            // Folder no longer exists — clear the stale value
            saveFolderInput.value = '';
            saveAppSettings();
          }
          updateSidebarStatus();
        });
      }
    }

    if (chkOpenAfter  && typeof s.openAfter === 'boolean') chkOpenAfter.checked  = s.openAfter;
    if (chkAddDate    && typeof s.addDate   === 'boolean') chkAddDate.checked    = s.addDate;

    // Восстанавливаем состояние чекбоксов шаблонов
    if (s.tplChecked && typeof s.tplChecked === 'object') {
      document.querySelectorAll('.tpl-item[data-template] input[type="checkbox"]').forEach((cb) => {
        const key = cb.closest('.tpl-item')?.dataset.template;
        if (key && typeof s.tplChecked[key] === 'boolean') {
          cb.checked = s.tplChecked[key];
        }
      });
    }
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
  // Clear valid indicator whenever value is wiped programmatically
  if (!el.value) el.classList.remove('input-valid');
}

function clearAllInputs() {
  Object.values(FIELD_MAP).forEach((id) => setInputValue(id, ''));
}

/** Mark all non-empty, non-readonly tracked fields as visually valid. */
function refreshValidStates() {
  for (const id of FIELD_IDS) {
    const el = document.getElementById(id);
    if (!el || el.readOnly) continue;
    el.classList.toggle('input-valid', el.value.trim() !== '');
  }
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

function onInputChange(el, currentValue) {
  if (!el) return;
  const inputId  = el.id;
  const original = originalValues[inputId] ?? '';
  const current  = currentValue.trim();
  if (current !== original) {
    dirtyInputIds.add(inputId);
    el.classList.add('input-dirty');
  } else {
    dirtyInputIds.delete(inputId);
    el.classList.remove('input-dirty');
  }
  updateDirtyState();
  // Remove valid indicator immediately when field is cleared while typing
  if (!current) el.classList.remove('input-valid');
  // Автосохранение: запускаем/сбрасываем дебаунс только если файл уже открыт
  if (currentFilePath) {
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(autoSave, 30000);
  }
  // Re-evaluate contract availability and tab status whenever owner fields change
  if (
    inputId.startsWith('owner1-') ||
    inputId.startsWith('owner2-') ||
    inputId.startsWith('owner3-')
  ) {
    updateContractAvailability();
    updateOwnerTabStatuses();
  }
  // Показываем/скрываем блок доверенности собственника
  if (inputId === 'owner1-Есть представитель' ||
      inputId === 'owner2-Есть представитель' ||
      inputId === 'owner3-Есть представитель') {
    const prefix = inputId.replace('-Есть представитель', '');
    applyOwnerPoaVisibility(prefix);
  }
  // Показываем/скрываем блок доверенности покупателя
  if (inputId === 'buyer-Есть представитель') {
    applyBuyerPoaVisibility();
  }
  // Re-evaluate object-type-dependent field visibility
  if (inputId === 'property-Тип объекта') {
    applyObjectTypeVisibility();
    autoUpdateCommission();
  }
  // Update block completion badge (дебаунс — не чаще раза в 200 мс)
  const dashIdx2 = inputId.indexOf('-');
  if (dashIdx2 !== -1) scheduleBlockCompletion(inputId.slice(0, dashIdx2));
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

/**
 * Снэпшот всех отслеживаемых полей за один проход по DOM.
 * Используется внутри buildPlaceholderData чтобы заменить 40+ getElementById-вызовов
 * одним проходом.
 */
function snapshotFields() {
  const snap = {};
  for (const id of FIELD_IDS) {
    const el = document.getElementById(id);
    if (el) snap[id] = el.value.trim();
  }
  return snap;
}

document.getElementById('deal-body').addEventListener('input', (e) => {
  const id = e.target.id;
  if (!id || !FIELD_IDS.has(id)) return;
  if (e.target.dataset.numeric) applyNumericFormat(e.target);
  onInputChange(e.target, e.target.value);
});

// Set/clear valid indicator when user leaves a field.
// Using focusout (bubbles) for event delegation.
document.getElementById('deal-body').addEventListener('focusout', (e) => {
  const el = e.target;
  if (!el.id || !FIELD_IDS.has(el.id) || el.readOnly) return;
  el.classList.toggle('input-valid', el.value.trim() !== '');
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
  const groups = { deal: {}, property: {}, owner1: {}, owner2: {}, owner3: {}, buyer: {} };
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
    await window.electronAPI.writeExcel(currentFileToken, currentFileToken, updates);
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
    await window.electronAPI.writeExcel(currentFileToken, currentFileToken, updates);
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

  let targetRef;
  try {
    targetRef = await window.electronAPI.saveFileDialog(defaultPath);
  } catch (err) {
    showToast('✖ Ошибка при открытии диалога: ' + err.message, 'error');
    return;
  }
  if (!targetRef) return;
  const targetPath = targetRef.path;
  const targetToken = targetRef.token;

  try {
    if (currentFilePath) {
      // Existing file loaded — copy its structure and write updates
      const updates = buildUpdates();
      await window.electronAPI.writeExcel(currentFileToken, targetToken, updates);
    } else {
      // No file loaded — build Excel from scratch using fields-config structure
      const fieldGroups = buildFieldGroups();
      const result = await window.electronAPI.createExcelFromData(fieldGroups, targetToken);
      if (!result.ok) throw new Error('Не удалось создать файл');
      // Bind the new rowMap so subsequent saves work via writeExcel normally
      rowMap = result.rowMap;
    }

    currentFilePath = targetPath;
    currentFileToken = targetToken;
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
      token: currentFileToken,
      label: baseName,
    });
  } catch (err) {
    showToast('✖ Не удалось сохранить файл: ' + err.message, 'error');
  }
}

// ============================================================
//  Close-app: main process asks renderer to save-then-close
// ============================================================
window.electronAPI?.onRequestSaveBeforeClose?.(async () => {
  await handleSave();
  window.electronAPI.closeApp();
});

// ============================================================
//  Owners count detection  (placed here — used by tab init below)
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

// ============================================================
//  Owner tabs — dynamic (1–3, add/remove)
// ============================================================
let ownerTabCount = 1; // how many owner tabs are currently visible

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn[data-tab]').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });
  document.querySelectorAll('.tab-pane').forEach((pane) => {
    pane.classList.toggle('active', pane.id === 'tab-pane-' + tabId);
  });
}

/** Show/hide tabs and controls to match ownerTabCount. */
function updateOwnerTabs() {
  const keys = ['owner1', 'owner2', 'owner3'];
  keys.forEach((key, i) => {
    const btn = document.querySelector(`.tab-btn[data-tab="${key}"]`);
    if (btn) btn.hidden = ownerTabCount < (i + 1);
  });

  // × remove button: only on the last visible tab (owner2 or owner3)
  ['owner2', 'owner3'].forEach((key, i) => {
    const rem = document.querySelector(`.owner-tab-remove[data-remove="${key}"]`);
    if (rem) rem.hidden = ownerTabCount !== (i + 2);
  });

  // "+ Добавить" button: hidden when all 3 tabs are open
  const addBtn = document.getElementById('btn-add-owner');
  if (addBtn) addBtn.hidden = ownerTabCount >= 3;

  updateOwnerTabStatuses();
}

/** Refresh ✓/○ icons on each owner tab. */
function updateOwnerTabStatuses() {
  ['owner1', 'owner2', 'owner3'].forEach((key) => {
    const el = document.getElementById('tab-status-' + key);
    if (!el) return;
    const filled = isOwnerPresent(key);
    el.textContent = filled ? '✓' : '○';
    el.classList.toggle('owner-tab-status--filled', filled);
  });
}

/** Clear all form fields belonging to a given owner prefix. */
function clearOwnerFields(ownerPrefix) {
  for (const [mapKey, inputId] of Object.entries(FIELD_MAP)) {
    if (mapKey.startsWith(ownerPrefix + '-')) {
      setInputValue(inputId, '');
    }
  }
}

/** After loading Excel, reveal tabs for every filled owner. */
function syncOwnerTabsToData() {
  const filled = ['owner1', 'owner2', 'owner3'].filter(isOwnerPresent).length;
  ownerTabCount = Math.max(1, filled);
  updateOwnerTabs();
}

// "+ Добавить собственника" button
document.getElementById('btn-add-owner')?.addEventListener('click', () => {
  if (ownerTabCount >= 3) return;
  ownerTabCount++;
  updateOwnerTabs();
  switchTab('owner' + ownerTabCount);
});

// "×" remove buttons (event delegation on the tab-bar)
document.getElementById('owner-tab-bar')?.addEventListener('click', (e) => {
  const rem = e.target.closest('.owner-tab-remove');
  if (!rem) return;
  e.stopPropagation(); // don't trigger the parent tab-btn click
  const key = rem.dataset.remove; // 'owner2' or 'owner3'
  if (!key || key === 'owner1') return;
  clearOwnerFields(key);
  applyOwnerPoaVisibility(key);
  ownerTabCount = Math.max(1, ownerTabCount - 1);
  switchTab('owner' + ownerTabCount);
  updateOwnerTabs();
  refreshValidStates();
  updateBlockCompletion(null);
  updateContractAvailability();
  updateSidebarStatus();
});

// Clicks on tab buttons (delegated, skips remove-button clicks)
document.getElementById('owner-tab-bar')?.addEventListener('click', (e) => {
  const btn = e.target.closest('.tab-btn[data-tab]');
  if (!btn || e.target.closest('.owner-tab-remove')) return;
  switchTab(btn.dataset.tab);
});

// Initial state
updateOwnerTabs();

// ============================================================
//  Populate form from parsed data
// ============================================================
function populateForm(data) {
  clearAllInputs();
  dirtyInputIds.clear();
  originalValues = {};

  rowMap = data._rowMap || {};

  const blockKeys = ['deal', 'property', 'owner1', 'owner2', 'owner3', 'buyer'];
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
  refreshValidStates();
  btnSaveAs.disabled = false;
  syncOwnerTabsToData(); // reveal tabs for filled owners, update icons
  switchTab('owner1');
  updateContractAvailability();
  applyAllOwnerPoaVisibility();
  applyBuyerPoaVisibility();
  applyObjectTypeVisibility();
  // Уведомить UIController об обновлении формы
  document.dispatchEvent(new Event('form:populated'));
  updateBlockCompletion(null); // пересчитать все блоки после загрузки
}

// ============================================================
//  Clear form
// ============================================================
function confirmClearForm() {
  const confirmed = window.confirm('Очистить форму? Все заполненные поля будут сброшены.');
  if (!confirmed) return false;
  handleClearForm();
  return true;
}

function handleClearForm() {
  clearAllInputs();
  rowMap = {};
  originalValues = {};
  dirtyInputIds.clear();
  clearTimeout(autoSaveTimer); // не запускать автосохранение после очистки
  autoSaveTimer = null;
  currentFilePath = null;
  currentFileToken = null;
  filePathDisplay.value = '';
  fileSuccess.hidden = true;
  btnSave.disabled = true;
  btnSaveAs.disabled = true;
  if (commissionInput) commissionInput.value = '';
  setStatus('Готов к работе');
  ownerTabCount = 1;
  updateOwnerTabs();
  switchTab('owner1');
  resetContractAvailability();
  applyAllOwnerPoaVisibility();
  applyBuyerPoaVisibility();
  applyObjectTypeVisibility();
  // Уведомить UIController об очистке формы
  document.dispatchEvent(new Event('form:cleared'));
  updateBlockCompletion(null); // сбросить все бейджи после очистки
}

// ============================================================
//  Check data (basic validation)
// ============================================================

function handleCheckData() {
  // Единый валидатор (validation.js): покупатель всегда обязателен
  const issues  = window.Validator.getValidationIssues({ requireBuyer: true });
  const missing = issues.map(f => f.label);
  const seen    = new Set(issues.map(f => f.id));

  // Скрытые поля (display:none), которые isInputVisible пропускает,
  // но которые обязательны для генерации документов — проверяем напрямую
  function chkHidden(id, label) {
    if (seen.has(id)) return;
    const el = document.getElementById(id);
    if (el && isFieldEmpty(el.value)) missing.push(label);
  }
  chkHidden('deal-Количество собственников', 'Сделка → Кол-во собственников');
  chkHidden('deal-Ответственный риэлтер',    'Сделка → Риэлтер');

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

// Единый обработчик для bynInput: подписывается после загрузки тарифов,
// чтобы гарантировать корректный расчёт комиссии с первого символа
window.COMMISSION_CONFIG_READY.then(() => {
  if (bynInput) {
    bynInput.addEventListener('input', () => {
      autoUpdatePropis();
      autoUpdateCommission();
    });
    autoUpdatePropis();
    autoUpdateCommission(); // пересчитать, если значение уже введено
  }
});

// toInstrumental, detectGender, buildNameGenitive, buildNameDative
// и вспомогательные _decline* — перенесены в js/name-declension.js

/** Возвращает массив ключей заполненных собственников, например ['owner1', 'owner2']. */
function getActiveOwners() {
  return ['owner1', 'owner2', 'owner3'].filter(isOwnerPresent);
}

function getOwnersCount() {
  const count = getActiveOwners().length;
  return count > 0 ? count : 1;
}

// ============================================================
//  Block completion badges + glow animation
// ============================================================
const BLOCK_COMPLETE_LABELS = {
  deal:     'Сделка заполнена',
  property: 'Объект заполнен',
  owners:   'Собственники заполнены',
  buyer:    'Покупатель заполнен',
};

function getRequiredIssueCountByBlock() {
  if (!window.Validator) return null;
  return window.Validator.getValidationIssues({ requireBuyer: false }).reduce((acc, issue) => {
    acc[issue.block] = (acc[issue.block] || 0) + 1;
    return acc;
  }, {});
}

function hasBlockEditableData(blockId) {
  const block = document.getElementById(blockId);
  if (!block) return false;
  return [...block.querySelectorAll('input[type="text"]')].some((el) => {
    if (el.readOnly || !isInputVisible(el)) return false;
    if (el.getAttribute('aria-hidden') === 'true') return false; // skip hidden select proxies
    return !isFieldEmpty(el.value);
  });
}

function isBuyerRequiredByDeposit() {
  const depositBYN = (document.getElementById('deal-Сумма задатка BYN')?.value || '').trim();
  const depositUSD = (document.getElementById('deal-Сумма задатка USD')?.value || '').trim();
  return depositBYN !== '' || depositUSD !== '';
}

/** true если el виден в layout (нет скрытого предка) */
function isInputVisible(el) {
  // Inputs в скрытых tab-pane (display:none) имеют offsetParent === null.
  // Форм-инпуты никогда не имеют position:fixed, поэтому проверка getComputedStyle не нужна.
  if (el.offsetParent === null) return false;
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

/** Дебаунс-обёртка: откладывает updateBlockCompletion на 200 мс после последнего вызова.
 *  Таймер хранится per-prefix, чтобы быстрый ввод в разных секциях не сбрасывал друг друга. */
const _blockCompletionTimers = new Map();
function scheduleBlockCompletion(prefix) {
  clearTimeout(_blockCompletionTimers.get(prefix));
  _blockCompletionTimers.set(prefix, setTimeout(() => {
    _blockCompletionTimers.delete(prefix);
    updateBlockCompletion(prefix);
  }, 200));
}

/** Обновляет badge и класс ws-block--complete для нужных блоков */
function updateBlockCompletion(changedPrefix) {
  const toCheck = new Set();
  if (!changedPrefix) {
    ['deal', 'property', 'owners', 'buyer'].forEach(b => toCheck.add(b));
  } else if (changedPrefix.startsWith('owner')) {
    toCheck.add('owners');
  } else {
    toCheck.add(changedPrefix);
  }

  for (const blockId of toCheck) {
    const section = document.getElementById(`ws-${blockId}`);
    const badge   = document.getElementById(`blk-status-${blockId}`);
    if (!section || !badge) continue;

    const issueCountByBlock = getRequiredIssueCountByBlock();
    let complete;
    if (issueCountByBlock) {
      const hasRequiredIssues = (issueCountByBlock[`ws-${blockId}`] || 0) > 0;
      const hasData = hasBlockEditableData(`ws-${blockId}`);
      complete = !hasRequiredIssues && hasData;
      if (blockId === 'buyer' && !isBuyerRequiredByDeposit() && !hasBlockEditableData('ws-buyer')) complete = false;
    } else if (blockId === 'owners') {
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
      // Покупатель заполнен, но задаток не указан — предупреждение в заголовке блока
      const buyerWarnNoDeposit = blockId === 'buyer' && !isBuyerRequiredByDeposit();
      badge.textContent = buyerWarnNoDeposit ? 'Укажите сумму задатка' : (BLOCK_COMPLETE_LABELS[blockId] || '');
      badge.className   = buyerWarnNoDeposit ? 'ws-block-hdr-status ws-block-hdr-status--warn' : 'ws-block-hdr-status';
      if (!wasComplete) {
        // Перезапуск flash-анимации без принудительного reflow:
        // сбрасываем animationName на один кадр, затем возвращаем
        section.classList.remove('ws-block--complete-flash');
        section.style.animationName = 'none';
        requestAnimationFrame(() => {
          section.style.animationName = '';
          section.classList.add('ws-block--complete-flash');
        });
      }
    } else {
      section.classList.remove('ws-block--complete', 'ws-block--complete-flash');
      badge.textContent = '';
      badge.className   = 'ws-block-hdr-status';
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

// Кэшируем NodeList один раз при старте — итерируем по массивам вместо повторных querySelectorAll
const _tplOwnerItems  = [...document.querySelectorAll('.tpl-item[data-owners-required]')];
const _objectTypeEls  = [...document.querySelectorAll('[data-object-type]')];

function updateContractAvailability() {
  const count    = getOwnersCount();
  const tooltip  = OWNERS_TOOLTIP[count] || '';
  _tplOwnerItems.forEach((label) => {
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
  _tplOwnerItems.forEach((label) => {
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

  _objectTypeEls.forEach((el) => {
    const type = el.dataset.objectType;
    if (isEmpty) {
      el.style.display = 'none';
      return;
    }
    if (type === 'дом') {
      el.style.display = isHouse ? '' : 'none';
    } else if (type === 'квартира') {
      el.style.display = isFlat ? '' : 'none';
    } else if (type === 'жилое') {
      el.style.display = (isHouse || isFlat) ? '' : 'none';
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
  saveAppSettings();
}
function handleDeselectAll() {
  document.querySelectorAll('.tpl-item input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
  updateSidebarStatus();
  saveAppSettings();
}

// ============================================================
//  Sidebar status — count badge + warnings
// ============================================================
function updateSidebarStatus() {
  const badgeEl = document.getElementById('tpl-count-badge');
  if (!badgeEl) return;

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
  if (!e.target.matches('.tpl-item input[type="checkbox"]')) return;
  updateSidebarStatus();
  saveAppSettings();

  // Тост для шаблонов, генерирующих по одному файлу на собственника
  const MULTI_FILE_TEMPLATES = ['soglasie-obrabotka', 'doverennost-pnd'];
  if (e.target.checked) {
    const matchedTemplate = MULTI_FILE_TEMPLATES.find(
      key => e.target.closest(`.tpl-item[data-template="${key}"]`)
    );
    if (matchedTemplate) {
      const n    = getActiveOwners().length;
      const word = n === 1 ? 'файл' : n < 5 ? 'файла' : 'файлов';
      showToast(
        `ℹ Заполнено собственников: ${n} — будет сформировано ${n} ${word}`,
        'info',
        4500
      );
    }
  }
});

// React to manual edits / clear of the folder input
saveFolderInput.addEventListener('input', () => { saveFolderToken = null; updateSidebarStatus(); saveAppSettings(); });

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

async function loadExcelFile(fileRef) {
  const filePath = fileRef && typeof fileRef === 'object' ? fileRef.path : fileRef;
  const fileToken = fileRef && typeof fileRef === 'object' ? fileRef.token : null;
  if (!filePath || !fileToken) {
    showError('Файл должен быть выбран через диалог в текущей сессии.');
    return;
  }
  filePathDisplay.value = filePath;
  fileSuccess.hidden = true;
  setStatus('Чтение файла…');
  showLoader();

  let data;
  try {
    data = await window.electronAPI.readExcel(fileToken);
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
  currentFileToken = fileToken;
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
    token: fileToken,
    label: baseName,
  });
}

async function handleChooseFile() {
  hideError();
  let fileRef;
  try {
    fileRef = await window.electronAPI.openFileDialog();
  } catch (err) {
    showError('Не удалось открыть диалог выбора файла: ' + err.message);
    return;
  }
  if (!fileRef) return;
  await loadExcelFile(fileRef);
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

  setDropState(currentFilePath ? 'success' : 'idle');
  showError('Для безопасности откройте Excel-файл через кнопку выбора файла.');
});

document.getElementById('btn-drop-browse').addEventListener('click', handleChooseFile);
document.getElementById('btn-drop-change').addEventListener('click', handleChooseFile);

// ============================================================
//  Event listeners
// ============================================================
btnChooseFile.addEventListener('click', handleChooseFile);
btnSave.addEventListener('click', handleSave);
btnSaveAs.addEventListener('click', handleSaveAs);
btnClear.addEventListener('click', () => confirmClearForm());
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
  if (chosen) { saveFolderInput.value = chosen.path; saveFolderToken = chosen.token; updateSidebarStatus(); saveAppSettings(); }
});

// ============================================================
//  Template data helpers
// ============================================================
// Активный снэпшот полей (устанавливается на время buildPlaceholderData).
// Позволяет getField читать из кэша вместо DOM при массовом обходе.
let _currentSnap = null;

function getField(id) {
  if (_currentSnap) return _currentSnap[id] ?? '';
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
//  Плейсхолдеры в .docx: {{deal.number}}, {{owner1.fullName}} и т.д.
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

// _declinePatronymicDative, _declineFirstNameDative, _declineLastNameDative,
// _declinePatronymicGenitive, _declineFirstNameGenitive, _declineLastNameGenitive,
// detectGender, buildNameGenitive, buildNameDative — перенесены в js/name-declension.js

function findPoaOwnerBySelectValue(value, ownerBlocks) {
  const normalized = (value || '').trim();
  if (!normalized) return null;
  return ownerBlocks.find((owner, idx) => normalized === `${owner.fullName} (собственник №${idx + 1})`)
      || ownerBlocks.find(owner => owner.fullName === normalized)
      || null;
}

// ── Составной блок паспортных данных для Word-шаблонов ───────
// Пример результата: «Паспорт серии MP, номер 1234567, выдан
// Лидским РОВД 01.02.2020.»
function buildPassportDetails({ passportSeries, passportNumber, passportIssuedBy, passportIssueDate, enabled = true }) {
  if (!enabled) return '';

  const series = String(passportSeries || '').trim();
  const number = String(passportNumber || '').trim();
  const issuedBy = String(passportIssuedBy || '').trim();
  const issueDate = String(passportIssueDate || '').trim();
  const passportParts = [];

  if (series) passportParts.push(`паспорт серии ${series}`);
  else if (number) passportParts.push('Паспорт');
  if (number) passportParts.push(`номер ${number}`);

  const issuerParts = [issuedBy, issueDate].filter(Boolean).join(' ');
  if (issuerParts) passportParts.push(`выдан ${issuerParts}`);

  return passportParts.length ? `${passportParts.join(', ')}` : '';
}

// ── Данные представителя собственника (для Word-шаблонов) ────
function buildOwnerPoaBlock(ownerPrefix) {
  const p  = ownerPrefix + '-';
  const hasPoa = (getField(p + 'Есть представитель') || '').trim().toLowerCase() === 'да';

  const lastName   = getField(p + 'Представитель фамилия')   || '';
  const firstName  = getField(p + 'Представитель имя')        || '';
  const middleName = getField(p + 'Представитель отчество')   || '';
  const fullName   = [lastName, firstName, middleName].filter(Boolean).join(' ');
  const initials   = lastName && firstName
    ? lastName + ' ' + firstName[0] + '.' + (middleName ? middleName[0] + '.' : '')
    : fullName;

  const series = getField(p + 'Представитель паспорт серия') || '';
  const number = getField(p + 'Представитель паспорт номер') || '';

  const genitive = buildNameGenitive(lastName, firstName, middleName);
  const dative   = buildNameDative(lastName, firstName, middleName);
  const gender   = detectGender(middleName, firstName, lastName);

  return {
    hasPoa,
    lastName,
    firstName,
    middleName,
    fullName,
    initials,
    ...genitive,
    ...dative,
    passportSeries:   series,
    passportNumber:   number,
    passport:         [series, number].filter(Boolean).join(' '),
    id:               getField(p + 'Представитель идент. номер')  || '',
    passportIssueDate: getField(p + 'Представитель дата выдачи')  || '',
    passportIssuedBy:  getField(p + 'Представитель кем выдан')    || '',
    passportIssuedByInstrumental: toInstrumental(getField(p + 'Представитель кем выдан') || ''),
    passportDetails:   buildPassportDetails({
      passportSeries: series,
      passportNumber: number,
      passportIssuedBy: getField(p + 'Представитель кем выдан'),
      passportIssueDate: getField(p + 'Представитель дата выдачи'),
      enabled: hasPoa,
    }),
    address:           getField(p + 'Представитель адрес')         || '',
    poaNumber:         getField(p + 'Номер доверенности')          || '',
    poaDate:           getField(p + 'Дата доверенности')           || '',
    // «действующего» / «действующей» — определяется по отчеству/имени представителя
    poaAction: gender === 'f' ? 'действующей' : 'действующего',
  };
}

function applyOwnerPoaVisibility(ownerPrefix) {
  const val = (getField(ownerPrefix + '-Есть представитель') || '').trim().toLowerCase();
  const hasPoa = val === 'да';
  const tabPane = document.getElementById('tab-pane-' + ownerPrefix);
  if (!tabPane) return;
  const block = tabPane.querySelector('.owner-poa-block');
  if (block) block.hidden = !hasPoa;
}

function applyAllOwnerPoaVisibility() {
  applyOwnerPoaVisibility('owner1');
  applyOwnerPoaVisibility('owner2');
  applyOwnerPoaVisibility('owner3');
}

// ── Данные представителя покупателя (для Word-шаблонов) ──────
function buildBuyerPoaBlock() {
  const p = 'buyer-';
  const hasPoa = (getField(p + 'Есть представитель') || '').trim().toLowerCase() === 'да';

  const lastName   = getField(p + 'Представитель фамилия')   || '';
  const firstName  = getField(p + 'Представитель имя')        || '';
  const middleName = getField(p + 'Представитель отчество')   || '';
  const fullName   = [lastName, firstName, middleName].filter(Boolean).join(' ');
  const initials   = lastName && firstName
    ? lastName + ' ' + firstName[0] + '.' + (middleName ? middleName[0] + '.' : '')
    : fullName;

  const series = getField(p + 'Представитель паспорт серия') || '';
  const number = getField(p + 'Представитель паспорт номер') || '';

  const genitive = buildNameGenitive(lastName, firstName, middleName);
  const dative   = buildNameDative(lastName, firstName, middleName);
  const gender   = detectGender(middleName, firstName, lastName);

  return {
    hasPoa,
    lastName,
    firstName,
    middleName,
    fullName,
    initials,
    ...genitive,
    ...dative,
    passportSeries:   series,
    passportNumber:   number,
    passport:         [series, number].filter(Boolean).join(' '),
    id:               getField(p + 'Представитель идент. номер')  || '',
    passportIssueDate: getField(p + 'Представитель дата выдачи')  || '',
    passportIssuedBy:  getField(p + 'Представитель кем выдан')    || '',
    passportIssuedByInstrumental: toInstrumental(getField(p + 'Представитель кем выдан') || ''),
    passportDetails:   buildPassportDetails({
      passportSeries: series,
      passportNumber: number,
      passportIssuedBy: getField(p + 'Представитель кем выдан'),
      passportIssueDate: getField(p + 'Представитель дата выдачи'),
      enabled: hasPoa,
    }),
    address:           getField(p + 'Представитель адрес')         || '',
    poaNumber:         getField(p + 'Номер доверенности')          || '',
    poaDate:           getField(p + 'Дата доверенности')           || '',
    poaAction: gender === 'f' ? 'действующей' : 'действующего',
  };
}

function applyBuyerPoaVisibility() {
  const val = (getField('buyer-Есть представитель') || '').trim().toLowerCase();
  const hasPoa = val === 'да';
  const section = document.getElementById('ws-buyer');
  if (!section) return;
  const block = section.querySelector('.owner-poa-block');
  if (block) block.hidden = !hasPoa;
}

// Поля, входящие в блок персоны (порядок важен для ключа кэша)
const _PERSON_BLOCK_FIELDS = [
  'Фамилия', 'Имя', 'Отчество', 'Дата рождения',
  'Паспорт серия', 'Паспорт номер', 'Идентификационный номер',
  'Кем выдан', 'Дата выдачи', 'Адрес регистрации', 'Телефон',
];

// Кэш результатов buildPersonBlock: ключ = значения полей, поэтому
// естественно инвалидируется при любом изменении — явная очистка не нужна
const _personBlockCache = new Map();

function buildPersonBlock(prefix) {
  // Ключ = prefix + все значения полей через \x00 (символ, не встречающийся в данных)
  const cacheKey = prefix + _PERSON_BLOCK_FIELDS.map(f => getField(prefix + f)).join('\x00');
  if (_personBlockCache.has(cacheKey)) return _personBlockCache.get(cacheKey);

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
  const result = {
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
    passportDetails:              buildPassportDetails({
      passportSeries: series,
      passportNumber: number,
      passportIssuedBy: getField(prefix + 'Кем выдан'),
      passportIssueDate: getField(prefix + 'Дата выдачи'),
    }),
    address:                      getField(prefix + 'Адрес регистрации')      || '',
    phone:                        getField(prefix + 'Телефон')                || '',
    email:                        '',
  };
  _personBlockCache.set(cacheKey, result);
  return result;
}

function buildPlaceholderData() {
  // Один проход по DOM — все последующие getField() читают из снэпшота
  _currentSnap = snapshotFields();
  try {
  const propertyTypeRaw = (getField('property-Тип объекта') || '').trim().toLowerCase();
  const isHouse      = propertyTypeRaw === 'дом';
  const isApartment  = propertyTypeRaw === 'квартира';
  const isCommercial = propertyTypeRaw === 'коммерческая недвижимость';

  const _endDateRaw = getField('deal-Дата окончания договора') || '';
  const deal = {
    number:                    getField('deal-Номер договора')  || '',
    date:                      getField('deal-Дата договора')   || '',
    dateText:                  dateToLongRussian(getField('deal-Дата договора') || ''),
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

  const owner1 = { ...buildPersonBlock('owner1-'), share: getField('owner1-Доля собственности') || '', poa: buildOwnerPoaBlock('owner1') };
  const owner2 = { ...buildPersonBlock('owner2-'), share: getField('owner2-Доля собственности') || '', poa: buildOwnerPoaBlock('owner2') };
  const owner3 = { ...buildPersonBlock('owner3-'), share: getField('owner3-Доля собственности') || '', poa: buildOwnerPoaBlock('owner3') };

  // ── Вычисляемые поля подписанта: представитель если есть, иначе собственник ──
  // Используйте в шаблонах: {{owner1.signatoryInitials}}, {{owner1.signatoryFullName}} и т.д.
  function addSignatoryFields(owner) {
    const src = owner.poa.hasPoa ? owner.poa : owner;
    owner.signatoryInitials           = src.initials;
    owner.signatoryFullName           = src.fullName;
    owner.signatoryLastName           = src.lastName;
    owner.signatoryFirstName          = src.firstName;
    owner.signatoryMiddleName         = src.middleName;
    owner.signatoryPassport           = src.passport;
    owner.signatoryPassportSeries     = src.passportSeries;
    owner.signatoryPassportNumber     = src.passportNumber;
    owner.signatoryPassportIssueDate  = src.passportIssueDate;
    owner.signatoryPassportIssuedBy   = src.passportIssuedBy;
    owner.signatoryId                 = src.id;
    owner.signatoryAddress            = src.address;
    // Склонения
    owner.signatoryFullNameGenitive   = src.fullNameGenitive   || src.fullName;
    owner.signatoryFullNameDative     = src.fullNameDative     || src.fullName;
    owner.signatoryLastNameGenitive   = src.lastNameGenitive   || src.lastName;
    owner.signatoryLastNameDative     = src.lastNameDative     || src.lastName;
  }
  addSignatoryFields(owner1);
  addSignatoryFields(owner2);
  addSignatoryFields(owner3);

  // ── poaSuffix: готовая строка «, в лице ФИО, действующего согласно доверенности № X от Y» ──
  // Пустая строка если представителя нет.
  // Используйте в шаблонах: {{owner1.poaSuffix}}, {{owner2.poaSuffix}}, {{owner3.poaSuffix}}
  function addPoaSuffix(owner) {
    if (owner.poa && owner.poa.hasPoa) {
      const poa = owner.poa;
      const poaNameGen  = poa.fullNameGenitive || poa.fullName;
      const poaNumPart  = poa.poaNumber ? ` № ${poa.poaNumber}` : '';
      const poaDatePart = poa.poaDate   ? ` от ${poa.poaDate}`  : '';
      owner.poaSuffix      = `, в лице ${poaNameGen}, ${poa.poaAction} согласно доверенности${poaNumPart}${poaDatePart}`;
      const poaNumRef      = poa.poaNumber ? `№ ${poa.poaNumber}` : '';
      const poaDateRef     = poa.poaDate   ? `от ${poa.poaDate}`  : '';
      const poaRef         = [poaNumRef, poaDateRef].filter(Boolean).join(' ');
      owner.signatoryLabel = (poa.initials || poa.fullName) + (poaRef ? ` (доверенность ${poaRef})` : '');
    } else {
      owner.poaSuffix      = '';
      owner.signatoryLabel = owner.initials || owner.fullName || '';
    }
  }
  addPoaSuffix(owner1);
  addPoaSuffix(owner2);
  addPoaSuffix(owner3);

  const ownerBlocks = [owner1, owner2, owner3];
  const buyer  = { ...buildPersonBlock('buyer-'), poa: buildBuyerPoaBlock() };

  // ── Представитель покупателя: signatoryFields + poaSuffix ──
  addSignatoryFields(buyer);
  addPoaSuffix(buyer);

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
    (getNumericField('deal-Стоимость BYN') || '0').replace(',', '.')
  );
  const commissionCfg = isCommercial
    ? window.COMMISSION_CONFIG_COMMERCIAL
    : window.COMMISSION_CONFIG;
  const commissionResult = window.calculateCommission(
    priceBYNRaw,
    commissionCfg.baseValue,
    commissionCfg.brackets
  );
  // baseUnitsRounded: Math.round(priceBYN / baseValue) — только для акта выполненных работ
  const _baseUnitsRounded = commissionResult.baseUnits
    ? Math.round(commissionResult.baseUnits)
    : 0;
  // vatAmount: НДС 20% от суммы комиссии — только для акта выполненных работ
  const _vatAmountRaw = commissionResult.amountBYNRaw
    ? Math.round(commissionResult.amountBYNRaw * 20) / 100
    : 0;
  const _vatAmount = _vatAmountRaw
    ? (_vatAmountRaw % 1 === 0 ? String(_vatAmountRaw) : _vatAmountRaw.toFixed(2).replace(/\.?0+$/, ''))
    : '';
  // baseUnitsRoundedWords: целое число базовых величин прописью (ж.р. — «базовая величина»)
  const _baseUnitsRoundedWords = (_baseUnitsRounded && typeof window.integerToWords === 'function')
    ? window.integerToWords(_baseUnitsRounded, 'f')
    : '';
  // vatAmountWords: сумма НДС прописью в BYN
  const _vatAmountWords = (_vatAmountRaw && typeof window.moneyToText === 'function')
    ? window.moneyToText(String(_vatAmountRaw))
    : '';

  const commission = {
    percent:               commissionResult.percent != null ? String(commissionResult.percent) : '',
    amountBYN:             commissionResult.amountBYN  || '',
    amountWords:           commissionResult.amountWords || '',
    baseValue:             String(window.COMMISSION_CONFIG.baseValue),
    baseUnits:             commissionResult.baseUnits
                             ? commissionResult.baseUnits.toFixed(2).replace(/\.?0+$/, '')
                             : '',
    baseUnitsRounded:      _baseUnitsRounded ? String(_baseUnitsRounded) : '',
    baseUnitsRoundedWords: _baseUnitsRoundedWords,
    vatAmount:             _vatAmount,
    vatAmountWords:        _vatAmountWords,
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

  // ── Подписант (seller) — данные для преамбулы и блоков подписей ──
  // seller = данные того, кто фактически подписывает (представитель если есть, иначе собственник).
  // Используется в шаблоне: {{seller.fullName}}, {{seller.address}}, {{seller.initials}} и т.д.
  function buildSellerFromOwner(owner) {
    const src = (owner.poa && owner.poa.hasPoa) ? owner.poa : owner;
    return {
      fullName:                    src.fullName                    || '',
      initials:                    src.initials                    || '',
      lastName:                    src.lastName                    || '',
      firstName:                   src.firstName                   || '',
      middleName:                  src.middleName                  || '',
      address:                     src.address                     || '',
      passport:                    src.passport                    || '',
      passportSeries:              src.passportSeries              || '',
      passportNumber:              src.passportNumber              || '',
      passportIssuedBy:            src.passportIssuedBy            || '',
      passportIssuedByInstrumental: src.passportIssuedByInstrumental || '',
      passportIssueDate:           src.passportIssueDate           || '',
      id:                          src.id                          || '',
      phone:                       owner.phone                     || '', // телефон всегда от собственника
      poaNumber:                   (owner.poa && owner.poa.hasPoa) ? (owner.poa.poaNumber || '') : '',
      poaDate:                     (owner.poa && owner.poa.hasPoa) ? (owner.poa.poaDate   || '') : '',
    };
  }

  // signatories — массив подписантов для loop-блоков в шаблоне Договора_реклама.
  // Каждый элемент: { seller: {...} } — сохраняем ключ seller, чтобы {{seller.*}}
  // внутри цикла {#signatories} работали без изменения шаблона.
  const _activeOwners = [owner1, owner2, owner3].filter(o => o.fullName);
  const signatories = (_activeOwners.length > 0 ? _activeOwners : [owner1])
    .map(owner => ({ seller: buildSellerFromOwner(owner) }));

  // seller на корневом уровне — для преамбулы договора (первый собственник / представитель)
  const seller = signatories[0].seller;

  // ── Количество экземпляров ──────────────────────────────────
  // Зависит от числа заполненных собственников: каждый собственник
  // получает свой экземпляр, плюс один экземпляр для исполнителя.
  const _filledOwnerCount = [owner1, owner2, owner3].filter(o => o.fullName).length || 1;
  const copies = {
    count:      _filledOwnerCount + 1,
    countWords: ['двух', 'трёх', 'четырёх'][_filledOwnerCount - 1] || String(_filledOwnerCount + 1),
  };

  // ── Рекламный пакет — прейскурант ──────────────────────────
  // Устанавливается через модалку перед генерацией Договора_реклама.
  // 0 означает «не выбран» — все галочки пустые.
  const _pkg = window._reklamaPackage || 0;
  const cb_1 = _pkg === 1 ? '☑' : '☐';
  const cb_2 = _pkg === 2 ? '☑' : '☐';
  const cb_3 = _pkg === 3 ? '☑' : '☐';
  const cb_4 = _pkg === 4 ? '☑' : '☐';

  // person — первый заполненный собственник (для шаблонов {{person.*}}:
  // soglasie-obrabotka, doverennost-pnd). При генерации generate-функции
  // переопределяют person конкретным собственником через { ...baseData, person }.
  const person = [owner1, owner2, owner3].find(o => o.fullName) || owner1;

  return { deal, property, owner1, owner2, owner3, buyer, agent, agency, keys, money, commission, deposit, copies, cb_1, cb_2, cb_3, cb_4, seller, signatories, person };
  } finally {
    _currentSnap = null;
  }
}


// ============================================================
// Вспомогательная функция: создаёт метод generate для заданного ключа шаблона
function makeGenerate(key) {
  return async function(outputDir, options) {
    return window.electronAPI.generateDocument(key, buildPlaceholderData(), outputDir, options);
  };
}

const TEMPLATE_REGISTRY = {
  'doverennost-pnd': {
    label: 'Доверенность ПНД',
    // Генерируем отдельный файл на каждого заполненного собственника
    generate: async function(outputDir, options) {
      const baseData = buildPlaceholderData();
      const ownerKeys = getActiveOwners();
      const results = [];
      for (const ownerKey of ownerKeys) {
        const owner = baseData[ownerKey];

        // poaSuffix и signatoryLabel — аналогично согласию на обработку данных
        let poaSuffix = '';
        if (owner.poa && owner.poa.hasPoa) {
          const poa = owner.poa;
          const poaNameGen  = poa.fullNameGenitive || poa.fullName;
          const poaNumPart  = poa.poaNumber ? ` № ${poa.poaNumber}` : '';
          const poaDatePart = poa.poaDate   ? ` от ${poa.poaDate}`  : '';
          poaSuffix = `, в лице ${poaNameGen}, ${poa.poaAction} согласно доверенности${poaNumPart}${poaDatePart}`;
        }

        let signatoryLabel = owner.initials || owner.fullName || '';
        if (owner.poa && owner.poa.hasPoa) {
          const poa = owner.poa;
          const poaNumPart  = poa.poaNumber ? `№ ${poa.poaNumber}` : '';
          const poaDatePart = poa.poaDate   ? `от ${poa.poaDate}`  : '';
          const poaRef      = [poaNumPart, poaDatePart].filter(Boolean).join(' ');
          signatoryLabel    = (poa.initials || poa.fullName) + (poaRef ? ` (доверенность ${poaRef})` : '');
        }

        // person — данные конкретного собственника для шаблона {{person.*}}
        const person = { ...owner, poaSuffix, signatoryLabel };
        const data   = { ...baseData, person };

        // Имя файла берётся из фамилии самого собственника (не представителя)
        const lastName = owner.lastName || owner.fullName || ownerKey;
        const outName  = `Доверенность ПНД — ${lastName}.docx`;
        const result = await window.electronAPI.generateDocument(
          'doverennost-pnd', data, outputDir, { ...options, outOverride: outName }
        );
        results.push({ result, outName, label: `Доверенность ПНД — ${owner.initials || lastName}` });
      }
      return { _multiFile: true, results };
    },
  },
  'raspiska-klyuchi':   { label: 'Расписка в получении ключей',                                                          generate: makeGenerate('raspiska-klyuchi') },
  'reklama':            { label: 'Договор на оказание рекламных услуг',                                                  generate: makeGenerate('reklama') },
  'rastorzhenie':       { label: 'Соглашение о расторжении',                                                             generate: makeGenerate('rastorzhenie') },
  'zapros-pnd':         { label: 'Запрос на ПНД',                                                                        generate: makeGenerate('zapros-pnd') },
  'zapros-rsc':         { label: 'Запрос в РСЦ (Справка)',                                                              generate: makeGenerate('zapros-rsc') },
  'zapros-rsc-privat':  { label: 'Запрос в РСЦ (Приватизация)',                                                          generate: makeGenerate('zapros-rsc-privat') },
  'soglasie-obrabotka': {
    label: 'Согласие на обработку данных',
    // Генерируем отдельный файл на каждого заполненного собственника
    generate: async function(outputDir, options) {
      const baseData = buildPlaceholderData();
      const ownerKeys = getActiveOwners();
      const results = [];
      for (const ownerKey of ownerKeys) {
        const owner = baseData[ownerKey];

        // poaSuffix:
        //   Без представителя: '' (пустая строка)
        //   С представителем:  ', в лице Ковалёвой Ирины Сергеевны, действующей
        //                        согласно доверенности № 12 от 01.01.2024'
        //
        // Шаблон в Word:
        //   Я, {{person.fullName}}, дата рождения: {{person.birthDate}},
        //   идентификационный номер паспорта {{person.id}}{{person.poaSuffix}}
        //   в соответствии со статьёй ...
        let poaSuffix = '';
        if (owner.poa && owner.poa.hasPoa) {
          const poa = owner.poa;
          // Имя представителя в родительном падеже («Ковалёвой Ирины Сергеевны»)
          const poaNameGen  = poa.fullNameGenitive || poa.fullName;
          const poaNumPart  = poa.poaNumber ? ` № ${poa.poaNumber}` : '';
          const poaDatePart = poa.poaDate   ? ` от ${poa.poaDate}`  : '';
          poaSuffix =
            `, в лице ${poaNameGen}, ${poa.poaAction} согласно доверенности${poaNumPart}${poaDatePart}`;
        }

        // signatoryLabel — строка для строки подписи:
        //   Без представителя: «Козловский Андрей Николаевич»
        //   С представителем:  «Ковалёва Ирина Сергеевна (доверенность № 1-2345 от 10.01.2026)»
        let signatoryLabel = owner.initials || owner.fullName || '';
        if (owner.poa && owner.poa.hasPoa) {
          const poa = owner.poa;
          const poaNumPart  = poa.poaNumber ? `№ ${poa.poaNumber}` : '';
          const poaDatePart = poa.poaDate   ? `от ${poa.poaDate}`  : '';
          const poaRef      = [poaNumPart, poaDatePart].filter(Boolean).join(' ');
          signatoryLabel    = (poa.initials || poa.fullName) + (poaRef ? ` (доверенность ${poaRef})` : '');
        }

        // person — данные конкретного собственника для шаблона {{person.*}}
        const person = { ...owner, poaSuffix, signatoryLabel };
        const data   = { ...baseData, person };

        // Имя файла всегда берётся из фамилии самого собственника (не представителя)
        const lastName = owner.lastName || owner.fullName || ownerKey;
        const outName  = `Согласие на обработку данных — ${lastName}.docx`;
        const result = await window.electronAPI.generateDocument(
          'soglasie-obrabotka', data, outputDir, { ...options, outOverride: outName }
        );
        results.push({ result, outName, label: `Согласие — ${owner.initials || lastName}` });
      }
      return { _multiFile: true, results };
    },
  },
  'dkp-3-eksklyuziv':  { label: 'Договор оказания риэлтерских услуг ЭКС',                                               generate: makeGenerate('dkp-3-eksklyuziv') },
  'dkp-3-obshiy':      { label: 'Договор оказания риэлтерских услуг',                                                   generate: makeGenerate('dkp-3-obshiy') },
  'konvertaciya':      { label: 'Договор о конвертации валюты',                                                          generate: makeGenerate('konvertaciya') },
  'zadatok-standart':  { label: 'Договор задатка (стандартный)',                                                         generate: makeGenerate('zadatok-standart') },
  'dkp-fizlit-komstr':       { label: 'Договор оказания риэлтерских услуг (физическое лицо — коммерческая структура)', generate: makeGenerate('dkp-fizlit-komstr') },
  'akt-rielterskikh-uslug':  { label: 'Акт приёмки-сдачи оказанных риелтерских услуг',                                  generate: makeGenerate('akt-rielterskikh-uslug') },
  'akt-reklamnykh-uslug':    { label: 'Акт приёмки-сдачи рекламных услуг',                                               generate: makeGenerate('akt-reklamnykh-uslug') },
};

// ============================================================
//  Generate documents
// ============================================================
// ============================================================
//  Рекламный пакет — модальное окно выбора
// ============================================================
window._reklamaPackage = 0;

(function () {
  const overlay    = document.getElementById('reklama-overlay');
  const radios     = () => overlay.querySelectorAll('.reklama-radio');
  const btnConfirm = document.getElementById('reklama-confirm');
  const btnCancel  = document.getElementById('reklama-cancel');
  const btnClose   = document.getElementById('reklama-close');

  let _resolveModal = null;
  const modal = new ModalController(overlay, {
    initialFocus: '.reklama-radio, #reklama-cancel',
    onClose: () => resolveModal(false),
  });

  function resolveModal(confirmed) {
    if (_resolveModal) {
      _resolveModal(confirmed);
      _resolveModal = null;
    }
  }

  function openReklamaModal() {
    // Reset selection
    radios().forEach(r => { r.checked = false; });
    btnConfirm.disabled = true;
    modal.open();

    return new Promise((resolve) => {
      _resolveModal = resolve;
    });
  }

  function closeModal(confirmed) {
    resolveModal(confirmed);
    modal.close({ confirmed });
  }

  // Radio selection enables the Confirm button
  overlay.addEventListener('change', (e) => {
    if (e.target.classList.contains('reklama-radio')) {
      btnConfirm.disabled = false;
    }
  });

  btnConfirm.addEventListener('click', () => {
    const selected = overlay.querySelector('.reklama-radio:checked');
    if (!selected) return;
    window._reklamaPackage = parseInt(selected.value, 10);
    closeModal(true);
  });

  btnCancel.addEventListener('click',  () => closeModal(false));
  btnClose.addEventListener('click',   () => closeModal(false));

  // Expose for use in handleGenerate and btnPreview
  window._openReklamaModal = openReklamaModal;
}());

// ============================================================
//  Shared: execute generation for a list of template keys
// ============================================================
async function _executeGenerate(toGenerate, outputDir, options) {
  let successCount = 0;
  const errors = [];

  showLoader();
  let results;
  try {
    results = await Promise.allSettled(
      toGenerate.map(key => TEMPLATE_REGISTRY[key].generate(outputDir, options))
    );
  } finally {
    hideLoader();
  }

  results.forEach((r, i) => {
    const key   = toGenerate[i];
    const entry = TEMPLATE_REGISTRY[key];
    if (r.status === 'fulfilled') {
      const result = r.value;

      // ── Мульти-файловый результат (напр. Согласие на обработку данных) ──
      if (result && result._multiFile) {
        result.results.forEach(({ result: sub, label }) => {
          if (sub && sub.success) {
            successCount++;
            window.RecentDocs?.push({
              name:  sub.path.split(/[/\\]/).pop(),
              type:  'word',
              icon:  key,
              path:  sub.path,
              token: sub.token,
              label,
            });
            if (chkOpenAfter && chkOpenAfter.checked) {
              window.electronAPI.openFile(sub.token);
            }
          } else {
            errors.push(`${label}: ${sub?.error || 'неизвестная ошибка'}`);
          }
        });
        return;
      }

      if (result && result.success) {
        successCount++;
        window.RecentDocs?.push({
          name:  result.path.split(/[/\\]/).pop(),
          type:  'word',
          icon:  key,
          path:  result.path,
          token: result.token,
          label: entry.label,
        });
        if (chkOpenAfter && chkOpenAfter.checked) {
          window.electronAPI.openFile(result.token);
        }
      } else {
        errors.push(`${entry.label}: ${result?.error || 'неизвестная ошибка'}`);
      }
    } else {
      errors.push(`${entry.label}: ${r.reason?.message || 'неизвестная ошибка'}`);
    }
  });

  if (successCount > 0) {
    showToast(`✔ Сформировано: ${successCount} из ${toGenerate.length}`);
  }
  errors.forEach(msg => showToast(`✖ ${msg}`, 'error'));
}

btnGenerate.addEventListener('click', handleGenerate);

async function handleGenerate() {
  const outputDir = saveFolderToken;
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

  // If reklama is in the list → show package selection modal first
  if (toGenerate.includes('reklama')) {
    const confirmed = await window._openReklamaModal();
    if (!confirmed) return; // user cancelled
  }

  await _executeGenerate(toGenerate, outputDir, options);
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

const previewModal = new ModalController(previewOverlay, {
  initialFocus: '#preview-close',
  onClose: () => {
    previewTabs.innerHTML = '';
    previewContent.innerHTML = '';
  },
});

function closePreviewModal() {
  previewModal.close();
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
           ✖ Ошибка предпросмотра: ${escapeHtml(result?.error || 'неизвестная ошибка')}
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
      `<div style="padding:32px;color:var(--error-text);font-size:13px;">✖ ${escapeHtml(err.message)}</div>`;
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

  previewModal.open({ initialFocus: '#preview-close' });
  previewContent.innerHTML = '';
  loadPreviewTab(templateKeys[0], data);
}

previewCloseBtn.addEventListener('click', closePreviewModal);
previewCloseFooter.addEventListener('click', closePreviewModal);

btnPreview.addEventListener('click', async () => {
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

  // If reklama is in preview list → ask for package first
  if (toPreview.includes('reklama')) {
    const confirmed = await window._openReklamaModal();
    if (!confirmed) return;
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

  const modal = new ModalController(overlay, { initialFocus: '#about-close' });

  function openAbout()  { modal.open(); }
  function closeAbout() { modal.close(); }

  btnOpen.addEventListener('click', openAbout);
  btnClose.addEventListener('click', closeAbout);
  btnOk.addEventListener('click', closeAbout);

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

  const modal = new ModalController(overlay, {
    initialFocus: '#update-btn-confirm',
    shouldClose: (detail) => !detail?.reason || !stepNotify.hidden,
  });

  function openModal() { modal.open(); }
  function closeModal() { modal.close(); }

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
  applyAllOwnerPoaVisibility();
  applyBuyerPoaVisibility();
  applyObjectTypeVisibility();

  // ── Collapsible template groups ───────────────────────────
  (function initTplGroups() {
    const STORAGE_KEY = 'tplGroupsCollapsed';

    // Load saved collapsed state: { groupId: true/false }
    let collapsed = {};
    try { collapsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch {}

    function setGroupState(section, isCollapsed, animate) {
      const hdr = section.querySelector('.tpl-group-hdr');
      if (!animate) section.style.transition = 'none';
      section.classList.toggle('tpl-section--collapsed', isCollapsed);
      if (hdr) hdr.setAttribute('aria-expanded', String(!isCollapsed));
      if (!animate) {
        // force reflow then re-enable transitions
        section.offsetHeight; // eslint-disable-line no-unused-expressions
        section.style.transition = '';
      }
    }

    document.querySelectorAll('.tpl-section[data-group]').forEach(section => {
      const groupId = section.dataset.group;
      const hdr = section.querySelector('.tpl-group-hdr');
      if (!hdr) return;

      // Restore saved state without animation on first paint
      if (collapsed[groupId]) setGroupState(section, true, false);

      hdr.addEventListener('click', () => {
        const nowCollapsed = !section.classList.contains('tpl-section--collapsed');
        setGroupState(section, nowCollapsed, true);
        collapsed[groupId] = nowCollapsed;
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(collapsed)); } catch {}
      });

      // Keyboard: Space / Enter
      hdr.addEventListener('keydown', e => {
        if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); hdr.click(); }
      });
    });
  })();
}());
