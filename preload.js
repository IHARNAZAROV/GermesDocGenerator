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
  generateDoverennost: (data, outputDir, options) => ipcRenderer.invoke('word:generateDoverennost', data, outputDir, options),
  generateReklama:     (data, outputDir, options) => ipcRenderer.invoke('word:generateReklama',     data, outputDir, options),

  // Document generation — Расписка в получении ключей
  generateRaspiska:  (data, outputDir, options)   => ipcRenderer.invoke('word:generateRaspiska',  data, outputDir, options),

  // Document generation — Соглашение о расторжении
  generateRastorzhenie: (data, outputDir, options) => ipcRenderer.invoke('word:generateRastorzhenie', data, outputDir, options),

  // Document generation — Запрос на ПНД
  generateZaprosPnd: (data, outputDir, options)   => ipcRenderer.invoke('word:generateZaprosPnd', data, outputDir, options),

  // Document generation — Запрос в РСЦ
  generateZaprosRsc: (data, outputDir, options)   => ipcRenderer.invoke('word:generateZaprosRsc', data, outputDir, options),

  // Document generation — Согласие на обработку данных
  generateSoglasie: (data, outputDir, options)    => ipcRenderer.invoke('word:generateSoglasie', data, outputDir, options),

  // Document generation — Договор ЭКС 1 собственник
  generateDkp1Eksklyuziv: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp1Eksklyuziv', data, outputDir, options),

  // Document generation — Договор 1 собственник (общий)
  generateDkp1Obshiy: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp1Obshiy', data, outputDir, options),

  // Document generation — Договор о конвертации валюты
  generateKonvertaciya: (data, outputDir, options) => ipcRenderer.invoke('word:generateKonvertaciya', data, outputDir, options),

  // Document generation — Договор ЭКС (2 собственника, общий)
  generateDkp2Eksklyuziv: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp2Eksklyuziv', data, outputDir, options),

  // Document generation — Договор оказания риэлтерских услуг (2 собственника, общий)
  generateDkp2Obshiy: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp2Obshiy', data, outputDir, options),

  // Document generation — Договор задатка (стандартный)
  generateZadatokStandart: (data, outputDir, options) => ipcRenderer.invoke('word:generateZadatokStandart', data, outputDir, options),

  // Document generation — Договор физическое лицо — коммерческая структура
  generateDkpFizlitKomstr: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkpFizlitKomstr', data, outputDir, options),
  generateDkp3Eksklyuziv: (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp3Eksklyuziv', data, outputDir, options),
  generateDkp3Obshiy:     (data, outputDir, options) => ipcRenderer.invoke('word:generateDkp3Obshiy',     data, outputDir, options),

  // Template scanning — opens file dialog, rescans Excel, reloads window
  scanTemplate: ()                             => ipcRenderer.invoke('template:scan'),

  // Document preview — renders template in-memory, returns text paragraphs
  previewDocument: (templateKey, data)         => ipcRenderer.invoke('word:preview', templateKey, data),
});
