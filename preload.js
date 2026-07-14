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

  // Document generation
  generateDoverennost: (data)                  => ipcRenderer.invoke('word:generateDoverennost', data),
});
