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
  'zapros-rsc':         { tpl: 'Запрос_в_РСЦ.docx',                       out: 'Запрос в РСЦ (Справка).docx' },
  'zapros-rsc-privat':  { tpl: 'Запрос_в_РСЦ_приватизация.docx',          out: 'Запрос в РСЦ (Приватизация).docx' },
  'soglasie-obrabotka': { tpl: 'Согласие_на_обработку_данных.docx',       out: 'Согласие на обработку данных.docx' },
  'konvertaciya':       { tpl: 'Договор_о_конвертации.docx',               out: 'Договор о конвертации.docx' },
  'zadatok-standart':  { tpl: 'Договор_задатка.docx',                     out: 'Договор задатка.docx' },
  'dkp-3-obshiy':      { tpl: 'Договор_риэлтерских_услуг.docx',           out: 'Договор риэлтерских услуг.docx' },
  'dkp-3-eksklyuziv':  { tpl: 'Договор_риэлтерских_услуг_ЭКС.docx',      out: 'Договор риэлтерских услуг ЭКС.docx' },
  'dkp-fizlit-komstr':       { tpl: 'Договор_физ_лица_коммерция.docx',         out: 'Договор физ лица коммерция.docx' },
  'akt-rielterskikh-uslug':  { tpl: 'Акт_выполненных_работ.docx',              out: 'Акт выполненных работ.docx' },
  'akt-reklamnykh-uslug':    { tpl: 'Акт_приемки_рекламных_услуг.docx',       out: 'Акт приёмки-сдачи рекламных услуг.docx' },
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

const sessionPaths = new Map();

function createFileToken(filePath, kind) {
  const token = `${kind}:${Date.now()}:${Math.random().toString(36).slice(2)}`;
  sessionPaths.set(token, { path: filePath, kind });
  return token;
}

function getSessionPath(token, expectedKind) {
  if (typeof token !== 'string') throw new Error('Некорректный токен файла');
  const entry = sessionPaths.get(token);
  if (!entry || (expectedKind && entry.kind !== expectedKind)) {
    throw new Error('Путь не был выбран в текущей сессии');
  }
  return entry.path;
}

function ensureExcelPath(filePath) {
  const ext = path.extname(filePath || '').toLowerCase();
  if (ext !== '.xlsx' && ext !== '.xls') throw new Error('Разрешены только Excel-файлы');
  return filePath;
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
      sandbox: true,
    },
    icon: path.join(__dirname, 'assets', 'icon.ico'),
    show: false,
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    checkForUpdates(mainWindow);
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
  const filePath = ensureExcelPath(result.filePaths[0]);
  return { path: filePath, token: createFileToken(filePath, 'excel') };
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
  const filePath = ensureExcelPath(result.filePath);
  return { path: filePath, token: createFileToken(filePath, 'excel') };
});

// ============================================================
//  IPC — read Excel
// ============================================================
ipcMain.handle('excel:read', async (_event, fileToken) => {
  try {
    const filePath = getSessionPath(fileToken, 'excel');
    const ExcelReader = require('./excel/excel-reader');
    return ExcelReader.readFile(filePath);
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

// ============================================================
//  IPC — create Excel from scratch using fields-config structure
//  fieldGroups: { deal: {key: value}, property: {...}, ... }
//  Returns { ok: true, rowMap: { "block-fieldKey": rowNumber } }
// ============================================================
ipcMain.handle('excel:createFromData', async (_event, fieldGroups, targetToken) => {
  try {
    const targetPath = getSessionPath(targetToken, 'excel');
    const fieldsConfig = require('./fields-config.json');

    // Инвертируем BLOCK_HEADERS из excel-reader: { deal: 'СДЕЛКА', ... }
    const { BLOCK_HEADERS } = require('./excel/excel-reader');
    const BLOCK_HEADER_MAP = Object.fromEntries(
      Object.entries(BLOCK_HEADERS).map(([header, id]) => [id, header])
    );

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
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

// ============================================================
//  IPC — write Excel
//  updates: { [rowNumber]: value }  — only column B is touched
// ============================================================
ipcMain.handle('excel:write', async (_event, sourceToken, targetToken, updates) => {
  try {
    const sourcePath = getSessionPath(sourceToken, 'excel');
    const targetPath = getSessionPath(targetToken, 'excel');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourcePath);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) return { ok: false, error: 'Файл не содержит листов' };

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
  } catch (err) {
    return { ok: false, error: err.message };
  }
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
// Restore a previously saved folder path on startup (no dialog — just re-mint a token)
ipcMain.handle('folder:restore', (_event, folderPath) => {
  if (typeof folderPath !== 'string' || !folderPath) return null;
  try {
    const stat = fs.statSync(folderPath);
    if (!stat.isDirectory()) return null;
  } catch {
    return null; // folder no longer exists
  }
  return { path: folderPath, token: createFileToken(folderPath, 'folder') };
});

ipcMain.handle('dialog:selectFolder', async (_event, defaultPath) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Выбрать папку для сохранения',
    defaultPath: defaultPath || undefined,
    properties: ['openDirectory', 'createDirectory'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const folderPath = result.filePaths[0];
  return { path: folderPath, token: createFileToken(folderPath, 'folder') };
});

// ============================================================
//  IPC — open file with default OS application
// ============================================================
ipcMain.handle('shell:openExternal', async (_event, url) => {
  let parsed;
  try { parsed = new URL(url); } catch { return; }
  if (parsed.protocol !== 'https:' && parsed.protocol !== 'mailto:') return;
  await shell.openExternal(url);
});

ipcMain.handle('shell:openFile', async (_event, fileToken) => {
  try {
    // This channel is shared by generated Word files and recent Excel files.
    const filePath = getSessionPath(fileToken);
    if (!fs.existsSync(filePath)) {
      return { success: false, error: `Файл не найден: ${filePath}` };
    }

    const error = await shell.openPath(filePath);
    if (error) {
      return { success: false, error };
    }

    return { success: true, path: filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// ============================================================
//  IPC — preview document (render template, return text lines)
// ============================================================
ipcMain.handle('word:preview', async (_event, templateKey, data) => {
  const entry = TEMPLATE_FILES[templateKey];
  if (!entry) return { success: false, error: `Шаблон не найден: ${templateKey}` };
  try {
    const templatePath = path.join(__dirname, 'templates', 'working', entry.tpl);
    return previewWord(templatePath, data);
  } catch (err) {
    return { success: false, error: err.message };
  }
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
ipcMain.handle('word:generate', async (_event, templateKey, data, outputToken, options = {}) => {
  const entry = TEMPLATE_FILES[templateKey];
  if (!entry) return { success: false, error: `Неизвестный ключ шаблона: ${templateKey}` };
  const templatePath = path.join(__dirname, 'templates', 'working', entry.tpl);
  const resolvedDir  = getSessionPath(outputToken, 'folder');
  // outOverride allows the caller to specify a custom output filename (e.g. per-owner consent docs)
  const baseName     = options.outOverride || entry.out;
  const outputPath   = buildOutputPath(resolvedDir, baseName, options.addDate);
  if (!fs.existsSync(resolvedDir)) fs.mkdirSync(resolvedDir, { recursive: true });
  const result = await generateWord(templatePath, outputPath, data);
  if (result && result.success && result.path) {
    result.token = createFileToken(result.path, 'generated');
  }
  return result;
});