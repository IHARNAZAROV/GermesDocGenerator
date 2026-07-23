const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // File dialogs
  openFileDialog:  ()                          => ipcRenderer.invoke('dialog:openFile'),
  saveFileDialog:  (defaultPath)               => ipcRenderer.invoke('dialog:saveFile', defaultPath),

  // Excel read / write / create
  readExcel:       (filePath)                  => ipcRenderer.invoke('excel:read', filePath),
  writeExcel:      (sourcePath, targetPath, updates) =>
                     ipcRenderer.invoke('excel:write', sourcePath, targetPath, updates),
  createExcelFromData: (fieldGroups, targetPath) =>
                     ipcRenderer.invoke('excel:createFromData', fieldGroups, targetPath),

  // Dirty-state notification → main process
  notifyDirtyChange: (isDirty)                 => ipcRenderer.send('app:dirty-changed', isDirty),

  // Main → renderer: save before close request
  onRequestSaveBeforeClose: (callback)         => ipcRenderer.on('request-save-before-close', callback),

  // Renderer → main: confirm close after save
  closeApp: ()                                 => ipcRenderer.send('app:close-confirmed'),

  // Folder selection dialog
  selectFolder: (defaultPath)                  => ipcRenderer.invoke('dialog:selectFolder', defaultPath),

  // Open file with default OS app
  openFile: (filePath)                         => ipcRenderer.invoke('shell:openFile', filePath),

  // Document generation — единый универсальный метод
  generateDocument: (templateKey, data, outputDir, options) =>
    ipcRenderer.invoke('word:generate', templateKey, data, outputDir, options),

  // Template scanning — opens file dialog, rescans Excel, reloads window
  scanTemplate: ()                             => ipcRenderer.invoke('template:scan'),

  // Document preview — renders template in-memory, returns text paragraphs
  previewDocument: (templateKey, data)         => ipcRenderer.invoke('word:preview', templateKey, data),

  // Open external URL in the default browser
  openExternal: (url)                          => ipcRenderer.invoke('shell:openExternal', url),

  // Get absolute file path from a File object (drag-and-drop)
  getPathForFile: (file)                       => webUtils.getPathForFile(file),

  // ── Автообновление (Portable) ────────────────────────────────────────────
  // Main → Renderer: найдена новая версия { version, assetUrl, assetName }
  onUpdateAvailable: (callback) =>
    ipcRenderer.on('update-available', (_e, info) => callback(info)),

  // Main → Renderer: прогресс скачивания { percent }
  onUpdateDownloadProgress: (callback) =>
    ipcRenderer.on('update-download-progress', (_e, info) => callback(info)),

  // Main → Renderer: ошибка при обновлении { message }
  onUpdateError: (callback) =>
    ipcRenderer.on('update-error', (_e, info) => callback(info)),

  // Renderer → Main: пользователь подтвердил скачивание
  startUpdate: () => ipcRenderer.send('update-start-download'),
});
