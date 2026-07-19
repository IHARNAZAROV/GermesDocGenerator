/**
 * updater.js — Модуль автообновления для Portable-версии
 *
 * Логика:
 *  1. При старте запрашивает GitHub Releases API
 *  2. Сравнивает tag_name с текущей версией app.getVersion()
 *  3. Если есть обновление — уведомляет renderer (событие update-available)
 *  4. При согласии пользователя скачивает portable .exe с прогрессом
 *  5. Генерирует .bat (Windows) / .sh (POSIX) скрипт-замену
 *  6. Запускает скрипт detached и вызывает app.quit()
 */

'use strict';

const { app, ipcMain }  = require('electron');
const https             = require('https');
const fs                = require('fs');
const path              = require('path');
const os                = require('os');
const { spawn }         = require('child_process');

// ─── Настройки ──────────────────────────────────────────────────────────────
const GITHUB_OWNER = 'IHARNAZAROV';
const GITHUB_REPO  = 'GermesDocGenerator';
// ────────────────────────────────────────────────────────────────────────────

/**
 * Выбирает Portable-ассет из списка ассетов релиза.
 * Приоритет: файл с "portable" в имени → любой .exe без "setup" в имени.
 */
function findPortableAsset(assets) {
  if (!assets || assets.length === 0) return null;
  const exeAssets = assets.filter(a => a.name.toLowerCase().endsWith('.exe'));
  // 1. Ищем явно portable
  const portable = exeAssets.find(a => a.name.toLowerCase().includes('portable'));
  if (portable) return portable;
  // 2. Любой .exe, который не является installer/setup
  const standalone = exeAssets.find(a => !a.name.toLowerCase().includes('setup'));
  if (standalone) return standalone;
  // 3. Первый попавшийся .exe
  return exeAssets[0] || null;
}

/**
 * Делает HTTPS GET-запрос и возвращает тело ответа как строку.
 * Автоматически следует за одним редиректом (302/301).
 */
function httpsGet(url, options = {}) {
  return new Promise((resolve, reject) => {
    const opts = Object.assign({ headers: { 'User-Agent': 'GermesDocGenerator-Updater' } }, options);

    const request = (targetUrl) => {
      https.get(targetUrl, opts, (res) => {
        // Один редирект
        if ((res.statusCode === 301 || res.statusCode === 302) && res.headers.location) {
          return request(res.headers.location);
        }
        if (res.statusCode !== 200) {
          return reject(new Error(`HTTP ${res.statusCode} для ${targetUrl}`));
        }
        let data = '';
        res.on('data', chunk => { data += chunk; });
        res.on('end', () => resolve(data));
        res.on('error', reject);
      }).on('error', reject);
    };

    request(url);
  });
}

/**
 * Скачивает файл по URL в указанный путь, вызывая onProgress(percent) по ходу.
 * Следует за редиректами автоматически.
 */
function downloadFile(url, destPath, onProgress) {
  return new Promise((resolve, reject) => {
    const opts = { headers: { 'User-Agent': 'GermesDocGenerator-Updater' } };

    const request = (targetUrl) => {
      https.get(targetUrl, opts, (res) => {
        // Редирект (GitHub Assets всегда редиректит на S3)
        if ((res.statusCode === 301 || res.statusCode === 302) && res.headers.location) {
          return request(res.headers.location);
        }
        if (res.statusCode !== 200) {
          return reject(new Error(`HTTP ${res.statusCode} при скачивании`));
        }

        const total = parseInt(res.headers['content-length'] || '0', 10);
        let received = 0;
        const file = fs.createWriteStream(destPath);

        res.on('data', chunk => {
          received += chunk.length;
          file.write(chunk);
          if (total > 0 && typeof onProgress === 'function') {
            onProgress(Math.round((received / total) * 100));
          }
        });

        res.on('end', () => {
          file.end(() => resolve());
        });

        res.on('error', err => {
          file.destroy();
          reject(err);
        });
      }).on('error', reject);
    };

    request(url);
  });
}

/**
 * Сравнивает две версии в формате «1.2.3».
 * Возвращает true, если remoteVersion > localVersion.
 */
function isNewer(remoteVersion, localVersion) {
  const parse = v => v.replace(/^v/, '').split('.').map(Number);
  const [rMaj, rMin, rPatch] = parse(remoteVersion);
  const [lMaj, lMin, lPatch] = parse(localVersion);

  if (rMaj !== lMaj) return rMaj > lMaj;
  if (rMin !== lMin) return rMin > lMin;
  return rPatch > lPatch;
}

/**
 * Генерирует и записывает Windows .bat скрипт-замены в temp-директорию.
 * Скрипт: ждёт 2 с → перемещает новый exe → запускает его → удаляет себя.
 *
 * @param {string} newExePath    — путь к скачанному новому portable exe
 * @param {string} currentExePath — путь к текущему running exe
 * @returns {string} путь к созданному .bat файлу
 */
function writeBatScript(newExePath, currentExePath) {
  const scriptPath = path.join(os.tmpdir(), 'gg_update.bat');
  const content = `@echo off
chcp 65001 >nul

:: Ожидаем завершения Electron (5 секунд)
timeout /t 5 /nobreak >nul

:: Пробуем переместить файл — до 10 попыток (файл может быть ещё заблокирован)
set ATTEMPTS=0
:RETRY
set /a ATTEMPTS+=1
move /y "${newExePath}" "${currentExePath}"
if errorlevel 1 (
  if %ATTEMPTS% lss 10 (
    timeout /t 2 /nobreak >nul
    goto RETRY
  ) else (
    del "%~f0"
    exit /b 1
  )
)

:: Успех — запускаем обновлённое приложение
start "" "${currentExePath}"

:: Удаляем этот скрипт
del "%~f0"
`;
  fs.writeFileSync(scriptPath, content, { encoding: 'utf8' });
  return scriptPath;
}

/**
 * Генерирует и записывает POSIX .sh скрипт-замены в temp-директорию.
 *
 * @param {string} newExePath    — путь к скачанному новому файлу
 * @param {string} currentExePath — путь к текущему запущенному файлу
 * @returns {string} путь к созданному .sh файлу
 */
function writeShScript(newExePath, currentExePath) {
  const scriptPath = path.join(os.tmpdir(), 'gg_update.sh');
  const content = `#!/bin/sh
# Ожидаем завершения Electron (2 секунды)
sleep 2

# Перемещаем новый файл на место старого
mv -f "${newExePath}" "${currentExePath}"
chmod +x "${currentExePath}"

# Запускаем обновлённую программу
"${currentExePath}" &

# Удаляем этот скрипт
rm -f "$0"
`;
  fs.writeFileSync(scriptPath, content, { encoding: 'utf8', mode: 0o755 });
  return scriptPath;
}

/**
 * Основная функция проверки обновлений.
 * Вызывается из main.js после создания окна.
 *
 * @param {Electron.BrowserWindow} mainWindow
 */
async function checkForUpdates(mainWindow) {
  // Откладываем проверку на 3 секунды, чтобы не замедлять старт
  await new Promise(r => setTimeout(r, 3000));

  let latestRelease;
  try {
    const apiUrl = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`;
    const body   = await httpsGet(apiUrl);
    latestRelease = JSON.parse(body);
  } catch (err) {
    // Тихо — нет интернета или API недоступен
    console.log('[updater] Не удалось получить данные о релизе:', err.message);
    return;
  }

  const remoteVersion = (latestRelease.tag_name || '').replace(/^v/, '');
  const localVersion  = app.getVersion();

  if (!remoteVersion || !isNewer(remoteVersion, localVersion)) {
    console.log(`[updater] Обновлений нет. Текущая версия: ${localVersion}, последняя: ${remoteVersion}`);
    return;
  }

  // Ищем portable-ассет в релизе
  const assets   = latestRelease.assets || [];
  const asset    = findPortableAsset(assets);
  const assetUrl = asset ? asset.browser_download_url : null;

  if (!assetUrl) {
    console.log('[updater] Portable-ассет не найден в релизе.');
    return;
  }

  console.log(`[updater] Найдено обновление ${remoteVersion}. Ассет: ${asset.name}`);

  // Уведомляем renderer
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('update-available', {
      version:  remoteVersion,
      assetUrl,
      assetName: asset.name,
    });
  }

  // Слушаем согласие пользователя (один раз)
  ipcMain.once('update-start-download', () => {
    startDownloadAndReplace(mainWindow, assetUrl, asset.name, remoteVersion);
  });
}

/**
 * Скачивает ассет, создаёт скрипт-замену, запускает его и закрывает приложение.
 */
async function startDownloadAndReplace(mainWindow, assetUrl, assetName, newVersion) {
  const tempDir    = app.getPath('temp');
  const destPath   = path.join(tempDir, assetName);

  const sendProgress = (percent) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('update-download-progress', { percent });
    }
  };

  const sendError = (message) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('update-error', { message });
    }
  };

  try {
    console.log(`[updater] Начинаем загрузку ${assetName} → ${destPath}`);
    await downloadFile(assetUrl, destPath, sendProgress);
    console.log('[updater] Загрузка завершена.');
  } catch (err) {
    console.error('[updater] Ошибка загрузки:', err.message);
    sendError(`Ошибка загрузки обновления: ${err.message}`);
    return;
  }

  // Путь к текущему запущенному исполняемому файлу
  const currentExePath = process.execPath;

  let scriptPath;
  let spawnArgs;

  if (process.platform === 'win32') {
    scriptPath = writeBatScript(destPath, currentExePath);
    spawnArgs  = { shell: false, detached: true, stdio: 'ignore', windowsHide: true };
    spawn('cmd.exe', ['/c', scriptPath], spawnArgs).unref();
  } else {
    scriptPath = writeShScript(destPath, currentExePath);
    spawnArgs  = { shell: true, detached: true, stdio: 'ignore' };
    spawn('sh', [scriptPath], spawnArgs).unref();
  }

  console.log(`[updater] Скрипт замены запущен: ${scriptPath}. Закрываем приложение…`);

  // Небольшая задержка, чтобы скрипт успел стартовать
  setTimeout(() => {
    app.quit();
  }, 500);
}

module.exports = { checkForUpdates };
