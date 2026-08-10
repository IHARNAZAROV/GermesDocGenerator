'use strict';

const { contextBridge, ipcRenderer, webUtils } = require('electron');

// ─── Validators ──────────────────────────────────────────────────────────────

/** Нормализует значение до строки; бросает TypeError если не строка/число. */
function str(v, name) {
  if (typeof v !== 'string' && typeof v !== 'number')
    throw new TypeError(`electronAPI: "${name}" must be a string (got ${typeof v})`);
  return String(v).trim();
}

/** Безопасный вариант: строка или undefined → строка или undefined. */
function optStr(v, name) {
  if (v === undefined || v === null) return undefined;
  return str(v, name);
}

/** Простой объект (не массив, не null). */
function obj(v, name) {
  if (v === null || typeof v !== 'object' || Array.isArray(v))
    throw new TypeError(`electronAPI: "${name}" must be a plain object`);
  return v;
}

/** Массив. */
function arr(v, name) {
  if (!Array.isArray(v))
    throw new TypeError(`electronAPI: "${name}" must be an array`);
  return v;
}

/** Функция-обратного вызова. */
function fn(v, name) {
  if (typeof v !== 'function')
    throw new TypeError(`electronAPI: "${name}" must be a function`);
  return v;
}

/**
 * Подписывается на IPC-канал, оборачивает callback (убирает event),
 * возвращает функцию-отписки.
 */
function on(channel, callback) {
  fn(callback, 'callback');
  const listener = (_event, ...args) => callback(...args);
  ipcRenderer.on(channel, listener);
  return () => ipcRenderer.removeListener(channel, listener);
}

// ─── Exposed API ─────────────────────────────────────────────────────────────

contextBridge.exposeInMainWorld('electronAPI', {

  // File dialogs
  openFileDialog: () =>
    ipcRenderer.invoke('dialog:openFile'),

  saveFileDialog: (defaultPath) =>
    ipcRenderer.invoke('dialog:saveFile', optStr(defaultPath, 'defaultPath')),

  // Excel read / write / create
  readExcel: (filePath) =>
    ipcRenderer.invoke('excel:read', str(filePath, 'filePath')),

  writeExcel: (sourcePath, targetPath, updates) =>
    ipcRenderer.invoke(
      'excel:write',
      str(sourcePath, 'sourcePath'),
      str(targetPath, 'targetPath'),
      obj(updates, 'updates'),
    ),

  createExcelFromData: (fieldGroups, targetPath) =>
    ipcRenderer.invoke(
      'excel:createFromData',
      arr(fieldGroups, 'fieldGroups'),
      str(targetPath, 'targetPath'),
    ),

  // Dirty-state notification → main process
  notifyDirtyChange: (isDirty) => {
    if (typeof isDirty !== 'boolean')
      throw new TypeError('electronAPI: "isDirty" must be a boolean');
    ipcRenderer.send('app:dirty-changed', isDirty);
  },

  // Main → renderer: save before close request
  // Возвращает unsubscribe-функцию; callback не получает IPC event.
  onRequestSaveBeforeClose: (callback) =>
    on('request-save-before-close', callback),

  // Renderer → main: confirm close after save
  closeApp: () =>
    ipcRenderer.send('app:close-confirmed'),

  // Folder selection dialog
  selectFolder: (defaultPath) =>
    ipcRenderer.invoke('dialog:selectFolder', optStr(defaultPath, 'defaultPath')),

  // Restore a saved folder path on startup without re-opening the dialog
  restoreFolder: (folderPath) =>
    ipcRenderer.invoke('folder:restore', optStr(folderPath, 'folderPath')),

  // Open file with default OS app
  openFile: (filePath) =>
    ipcRenderer.invoke('shell:openFile', str(filePath, 'filePath')),

  // Document generation
  generateDocument: (templateKey, data, outputDir, options) =>
    ipcRenderer.invoke(
      'word:generate',
      str(templateKey, 'templateKey'),
      obj(data, 'data'),
      str(outputDir, 'outputDir'),
      options !== undefined ? obj(options, 'options') : undefined,
    ),

  // AI advertisement generation
  generateAd: (config, data) =>
    ipcRenderer.invoke(
      'ai:generateAd',
      obj(config, 'config'),
      obj(data, 'data'),
    ),

  // Template scanning
  scanTemplate: () =>
    ipcRenderer.invoke('template:scan'),

  // Document preview
  previewDocument: (templateKey, data) =>
    ipcRenderer.invoke(
      'word:preview',
      str(templateKey, 'templateKey'),
      obj(data, 'data'),
    ),

  // Open external URL in the default browser
  openExternal: (url) =>
    ipcRenderer.invoke('shell:openExternal', str(url, 'url')),

  // Get absolute file path from a File object (drag-and-drop)
  getPathForFile: (file) => {
    if (!(file instanceof File))
      throw new TypeError('electronAPI: "file" must be a File instance');
    return webUtils.getPathForFile(file);
  },

  // ── Автообновление (Portable) ─────────────────────────────────────────────
  // Все три метода возвращают unsubscribe-функцию; callback не получает IPC event.

  onUpdateAvailable: (callback) =>
    on('update-available', callback),

  onUpdateDownloadProgress: (callback) =>
    on('update-download-progress', callback),

  onUpdateError: (callback) =>
    on('update-error', callback),

  // Renderer → Main: пользователь подтвердил скачивание
  startUpdate: () =>
    ipcRenderer.send('update-start-download'),
});
