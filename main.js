const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs   = require('fs');
const ExcelJS = require('exceljs');
const { generateWord, previewWord } = require("./generator/word-generator");
const { checkForUpdates } = require('./updater');

// Map of template keys → { tpl: filename in templates/working/, out: output filename }
const TEMPLATE_FILES = {
  'doverennost-pnd':    { tpl: 'Доверенность_ПНД.docx',                  out: 'Доверенность ПНД.docx' },
  'raspiska-klyuchi':   { tpl: 'РАСПИСКА_в_получении_ключей.docx',        out: 'Расписка в получении ключей.docx' },
  'reklama':            { tpl: 'Договор_реклама.docx',                    out: 'Договор реклама.docx' },
  'rastorzhenie':       { tpl: 'Соглашение_о_расторжении.docx',           out: 'Соглашение о расторжении.docx' },
  'zapros-pnd':         { tpl: 'Запрос_на_ПНД.docx',                      out: 'Запрос на ПНД.docx' },
  'zapros-rsc':         { tpl: 'Запрос_в_РСЦ.docx',                       out: 'Запрос в РСЦ.docx' },
  'soglasie-obrabotka': { tpl: 'Согласие_на_обработку_данных.docx',       out: 'Согласие на обработку данных.docx' },
  'dkp-1-eksklyuziv':  { tpl: 'Договор_ЭКС_1_собств.docx',               out: 'Договор ЭКС 1 собств.docx' },
  'dkp-1-obshiy':      { tpl: 'Договор_1_собств_общий.docx',              out: 'Договор 1 собств общий.docx' },
  'konvertaciya':       { tpl: 'Договор_о_конвертации.docx',               out: 'Договор о конвертации.docx' },
  'zadatok-standart':  { tpl: 'Договор_задатка.docx',                     out: 'Договор задатка.docx' },
  'dkp-2-obshiy':      { tpl: 'Договор_2_собств_общий.docx',              out: 'Договор 2 собств общий.docx' },
  'dkp-2-eksklyuziv':  { tpl: 'Договор_ЭКС_2_собств.docx',               out: 'Договор ЭКС 2 собств.docx' },
  'dkp-3-obshiy':      { tpl: 'Договор_3_собств_общий.docx',              out: 'Договор 3 собств общий.docx' },
  'dkp-3-eksklyuziv':  { tpl: 'Договор_ЭКС_3_собств.docx',               out: 'Договор ЭКС 3 собств.docx' },
  'dkp-fizlit-komstr': { tpl: 'Договор_физ_лица_коммерция.docx',          out: 'Договор физ лица коммерция.docx' },
};

// ============================================================
//  Helper — build output file path, optionally appending date
// ============================================================
function buildOutputPath(dir, baseName, addDate) {
  if (!addDate) return path.join(dir, baseName);
  const now = new Date();
  const dd   = String(now.getDate()).padStart(2, '0');
  const mm   = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const ext  = path.extname(baseName);
  const name = path.basename(baseName, ext);
  return path.join(dir, `${name}_${dd}-${mm}-${yyyy}${ext}`);
}

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
    icon: path.join(__dirname, 'assets', 'icon.ico'),
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
  Menu.setApplicationMenu(null);
  createWindow();

  // === Автообновление Portable ===
  checkForUpdates(mainWindow);
  // ================================

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
//  IPC — create Excel from scratch using fields-config structure
//  fieldGroups: { deal: {key: value}, property: {...}, ... }
//  Returns { ok: true, rowMap: { "block-fieldKey": rowNumber } }
// ============================================================
ipcMain.handle('excel:createFromData', async (_event, fieldGroups, targetPath) => {
  const fieldsConfig = require('./fields-config.json');

  const BLOCK_HEADER_MAP = {
    'deal':     'СДЕЛКА',
    'property': 'ОБЪЕКТ',
    'seller':   'ПРОДАВЕЦ',
    'owner1':   'СОБСТВЕННИК №1',
    'owner2':   'СОБСТВЕННИК №2',
    'owner3':   'СОБСТВЕННИК №3',
    'buyer':    'ПОКУПАТЕЛЬ',
  };

  const workbook  = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Сделка');
  const rowMap    = {};
  let   rowNum    = 1;

  for (const group of fieldsConfig.groups) {
    const blockId = group.id;
    const header  = BLOCK_HEADER_MAP[blockId];
    if (!header) continue;

    // Block header row (column A only, no value in B)
    worksheet.getCell(rowNum, 1).value = header;
    rowNum++;

    for (const field of group.fields) {
      // Skip fields computed entirely by the app — they are never stored in Excel
      if (field.computed) continue;

      const value = ((fieldGroups || {})[blockId] || {})[field.key];
      worksheet.getCell(rowNum, 1).value = field.key;
      worksheet.getCell(rowNum, 2).value =
        (value !== undefined && value !== null && String(value).trim() !== '')
          ? String(value).trim()
          : null;

      rowMap[`${blockId}-${field.key}`] = rowNum;
      rowNum++;
    }
  }

  await workbook.xlsx.writeFile(targetPath);
  return { ok: true, rowMap };
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
//  IPC — select output folder dialog
// ============================================================
ipcMain.handle('dialog:selectFolder', async (_event, defaultPath) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Выбрать папку для сохранения',
    defaultPath: defaultPath || undefined,
    properties: ['openDirectory', 'createDirectory'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

// ============================================================
//  IPC — open file with default OS application
// ============================================================
ipcMain.handle('shell:openExternal', async (_event, url) => {
  await shell.openExternal(url);
});

ipcMain.handle('shell:openFile', async (_event, filePath) => {
  await shell.openPath(filePath);
});

// ============================================================
//  IPC — preview document (render template, return text lines)
// ============================================================
ipcMain.handle('word:preview', async (_event, templateKey, data) => {
  const entry = TEMPLATE_FILES[templateKey];
  if (!entry) return { success: false, error: `Шаблон не найден: ${templateKey}` };
  const templatePath = path.join(__dirname, 'templates', 'working', entry.tpl);
  return previewWord(templatePath, data);
});

// ============================================================
//  IPC — scan Excel template and update fields-config
// ============================================================
ipcMain.handle('template:scan', async () => {
  // 1. Диалог выбора Excel-шаблона
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Выбрать Excel-шаблон для обновления формы',
    filters: [{ name: 'Excel файлы', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return { canceled: true };

  const excelPath = result.filePaths[0];
  try {
    const { scanAndUpdate } = require('./excel/excel-scanner');
    const info = await scanAndUpdate(excelPath);
    // Перезагружаем рендерер чтобы подхватить новый js/fields-config.js
    mainWindow.reload();
    return { ok: true, ...info };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

// ============================================================
//  IPC — единый универсальный хендлер генерации Word-документов
// ============================================================
ipcMain.handle('word:generate', async (_event, templateKey, data, outputDir, options = {}) => {
  const entry = TEMPLATE_FILES[templateKey];
  if (!entry) return { success: false, error: `Неизвестный ключ шаблона: ${templateKey}` };
  const templatePath = path.join(__dirname, 'templates', 'working', entry.tpl);
  const resolvedDir  = outputDir || path.join(__dirname, 'output');
  const outputPath   = buildOutputPath(resolvedDir, entry.out, options.addDate);
  if (!fs.existsSync(resolvedDir)) fs.mkdirSync(resolvedDir, { recursive: true });
  return generateWord(templatePath, outputPath, data);
});