const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const { generateWord } = require("./generator/word-generator");

let mainWindow;
let isDirty = false; // renderer notifies us when dirty state changes

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1200,
    minHeight: 800,
    title: 'Генератор договоров ГермесГарант',
    backgroundColor: '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, 'assets', 'icon.png'),
    show: false,
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // Intercept close to handle unsaved changes
  mainWindow.on('close', (e) => {
    if (!isDirty) return; // nothing to ask

    e.preventDefault();

    dialog.showMessageBox(mainWindow, {
      type: 'question',
      title: 'Несохраненные изменения',
      message: 'Есть несохраненные изменения. Что сделать?',
      buttons: ['💾 Сохранить', '❌ Не сохранять', '↩ Отмена'],
      defaultId: 0,
      cancelId: 2,
      noLink: true,
    }).then(({ response }) => {
      if (response === 0) {
        // Ask renderer to save, then close
        mainWindow.webContents.send('request-save-before-close');
      } else if (response === 1) {
        isDirty = false;
        mainWindow.destroy();
      }
      // response === 2: cancel — do nothing
    });
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ============================================================
//  IPC — file open dialog
// ============================================================
ipcMain.handle('dialog:openFile', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Выбрать Excel файл',
    filters: [{ name: 'Excel файлы', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

// ============================================================
//  IPC — file save dialog
// ============================================================
ipcMain.handle('dialog:saveFile', async (_event, defaultPath) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'Сохранить как…',
    defaultPath,
    filters: [{ name: 'Excel файлы', extensions: ['xlsx'] }],
  });
  if (result.canceled || !result.filePath) return null;
  return result.filePath;
});

// ============================================================
//  IPC — read Excel
// ============================================================
ipcMain.handle('excel:read', async (_event, filePath) => {
  const ExcelReader = require('./excel/excel-reader');
  return ExcelReader.readFile(filePath);
});

// ============================================================
//  IPC — write Excel
//  updates: { [rowNumber]: value }  — only column B is touched
// ============================================================
ipcMain.handle('excel:write', async (_event, sourcePath, targetPath, updates) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourcePath);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('Файл не содержит листов');

  for (const [rowNumStr, value] of Object.entries(updates)) {
    const rowNum = parseInt(rowNumStr, 10);
    const cell = worksheet.getCell(rowNum, 2); // column B
    // Preserve cell style — only overwrite the value
    cell.value = (value !== null && value !== undefined && String(value).trim() !== '')
      ? String(value).trim()
      : null;
  }

  await workbook.xlsx.writeFile(targetPath);
  return { ok: true };
});

// ============================================================
//  IPC — dirty state notifications from renderer
// ============================================================
ipcMain.on('app:dirty-changed', (_event, dirty) => {
  isDirty = !!dirty;
});

// ============================================================
//  IPC — renderer confirmed save-and-close
// ============================================================
ipcMain.on('app:close-confirmed', () => {
  isDirty = false;
  if (mainWindow) mainWindow.destroy();
});

// ============================================================
//  IPC — generate "Доверенность ПНД" from current form data
// ============================================================
ipcMain.handle('word:generateDoverennost', async (_event, data) => {
  const fs = require('fs');

  const templatePath = path.join(__dirname, 'templates', 'working', 'Доверенность_ПНД.docx');
  const outputDir    = path.join(__dirname, 'output');
  const outputPath   = path.join(outputDir, 'Доверенность ПНД.docx');

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  return generateWord(templatePath, outputPath, data);
});