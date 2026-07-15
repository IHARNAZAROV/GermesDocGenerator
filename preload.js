const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // File dialogs
  openFileDialog:  ()                          => ipcRenderer.invoke('dialog:openFile'),
  saveFileDialog:  (defaultPath)               => ipcRenderer.invoke('dialog:saveFile', defaultPath),

  // Excel read / write
  readExcel:       (filePath)                  => ipcRenderer.invoke('excel:read', filePath),
  writeExcel:      (sourcePath, targetPath, updates) =>
                     ipcRenderer.invoke('excel:write', sourcePath, targetPath, updates),

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

  // Document generation
  generateDoverennost: (data, outputDir)       => ipcRenderer.invoke('word:generateDoverennost', data, outputDir),
  generateReklama:     (data, outputDir)       => ipcRenderer.invoke('word:generateReklama',     data, outputDir),

  // Document generation — Расписка в получении ключей
  generateRaspiska:  (data, outputDir)         => ipcRenderer.invoke('word:generateRaspiska',  data, outputDir),

  // Document generation — Запрос в РСЦ
  generateZaprosRsc: (data, outputDir)         => ipcRenderer.invoke('word:generateZaprosRsc', data, outputDir),

  // Template scanning — opens file dialog, rescans Excel, reloads window
  scanTemplate: ()                             => ipcRenderer.invoke('template:scan'),
});
