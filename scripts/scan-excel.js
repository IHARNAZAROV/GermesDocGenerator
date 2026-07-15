#!/usr/bin/env node
/**
 * scan-excel.js — сканер шаблона Excel
 *
 * Использование:
 *   node scripts/scan-excel.js <путь/к/шаблону.xlsx>
 *
 * Что делает:
 *   1. Читает все листы указанного Excel-файла.
 *   2. Находит все блоки (СДЕЛКА, ОБЪЕКТ, …) и поля внутри них.
 *   3. Сливает результат с существующим fields-config.json:
 *      — новые поля добавляются с дефолтными настройками;
 *      — существующие поля сохраняют вручную заданные label/type/section/… ;
 *      — поля, исчезнувшие из Excel, удаляются (кроме полей с "computed": true).
 *   4. Перезаписывает:
 *      — fields-config.json   (источник истины для Node.js / скрипта)
 *      — js/fields-config.js  (браузерная версия для Electron-рендерера)
 */

'use strict';

const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');

// ── Пути ──────────────────────────────────────────────────────
const ROOT       = path.resolve(__dirname, '..');
const CONFIG_JSON = path.join(ROOT, 'fields-config.json');
const CONFIG_JS   = path.join(ROOT, 'js', 'fields-config.js');

// ── Аргументы ─────────────────────────────────────────────────
const excelPath = process.argv[2];
if (!excelPath) {
  console.error('Использование: node scripts/scan-excel.js <путь/к/шаблону.xlsx>');
  process.exit(1);
}
const resolvedExcel = path.resolve(excelPath);
if (!fs.existsSync(resolvedExcel)) {
  console.error(`Файл не найден: ${resolvedExcel}`);
  process.exit(1);
}

// ── Утилиты ───────────────────────────────────────────────────
function cellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) return v.toLocaleDateString('ru-RU');
    if (v.result  !== undefined) return String(v.result);
    if (v.text    !== undefined) return String(v.text);
    if (v.richText !== undefined) return v.richText.map(r => r.text || '').join('');
    return String(v);
  }
  return String(v);
}

/** Инфер типа поля по имени */
function inferType(key) {
  const k = key.toLowerCase();
  if (k.startsWith('дата')) return 'date';
  return 'text';
}

/** Инфер метки из ключа (добавить двоеточие) */
function inferLabel(key) {
  return key + ':';
}

// ── Сканирование Excel ────────────────────────────────────────
async function scanExcel(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  // Загружаем существующий конфиг (для слияния)
  let existing = { meta: {}, groups: [] };
  if (fs.existsSync(CONFIG_JSON)) {
    try { existing = JSON.parse(fs.readFileSync(CONFIG_JSON, 'utf8')); }
    catch (e) { console.warn('Не удалось прочитать fields-config.json, начинаем заново.'); }
  }

  // Индекс существующих групп и полей для быстрого поиска
  const existingGroups = {};
  for (const g of (existing.groups || [])) {
    existingGroups[g.id] = {};
    for (const f of (g.fields || [])) {
      existingGroups[g.id][f.key] = f;
    }
  }

  // Построим карту excelHeader → groupId из существующего конфига
  const headerToGroupId = {};
  for (const g of (existing.groups || [])) {
    if (g.excelHeader) headerToGroupId[g.excelHeader] = g.id;
  }

  // Сканируем каждый лист
  const scannedGroups = {}; // groupId → { excelHeader, sheet, fields: [{ key, row }] }

  wb.eachSheet((ws) => {
    let currentGroupId   = null;
    let currentGroupHdr  = null;

    ws.eachRow((row, rowNum) => {
      const a = cellText(row.getCell(1)).trim();
      if (!a) return;

      // Проверяем, является ли строка заголовком блока
      const groupId = headerToGroupId[a];
      if (groupId !== undefined) {
        currentGroupId  = groupId;
        currentGroupHdr = a;
        if (!scannedGroups[currentGroupId]) {
          scannedGroups[currentGroupId] = {
            excelHeader: currentGroupHdr,
            sheet: ws.name,
            fields: [],
          };
        }
        return;
      }

      if (currentGroupId) {
        // Это поле внутри блока
        scannedGroups[currentGroupId].fields.push({ key: a, row: rowNum });
      }
    });
  });

  // ── Слияние: строим итоговый список групп ─────────────────
  const mergedGroups = [];

  // Сохраняем порядок и структуру существующего конфига
  for (const existingGroup of (existing.groups || [])) {
    const gid     = existingGroup.id;
    const scanned = scannedGroups[gid];
    const oldByKey = existingGroups[gid] || {};

    if (!scanned && Object.values(oldByKey).every(f => !f.computed)) {
      // Группа исчезла из Excel полностью, пропускаем
      console.log(`  [удалена группа] ${gid}`);
      continue;
    }

    const scannedKeys  = new Set((scanned?.fields || []).map(f => f.key));
    const mergedFields = [];

    // 1. Поля, которые уже были в конфиге
    for (const oldField of (existingGroup.fields || [])) {
      if (oldField.computed) {
        // Вычисляемые поля (не из Excel) всегда сохраняем
        mergedFields.push(oldField);
        continue;
      }
      if (scannedKeys.has(oldField.key)) {
        // Поле ещё есть в Excel — сохраняем всю метаинформацию
        mergedFields.push(oldField);
        scannedKeys.delete(oldField.key); // отмечаем как обработанное
      } else {
        console.log(`  [удалено поле] ${gid}.${oldField.key}`);
        // Удаляем: поля нет в Excel
      }
    }

    // 2. Новые поля, найденные в Excel, которых ещё не было в конфиге
    if (scanned) {
      for (const { key } of scanned.fields) {
        if (!scannedKeys.has(key)) continue; // уже обработано выше
        console.log(`  [новое поле] ${gid}.${key}`);
        mergedFields.push({
          key,
          label: inferLabel(key),
          type:  inferType(key),
        });
      }
    }

    mergedGroups.push({
      ...existingGroup,
      fields: mergedFields,
    });
  }

  // ── Мета ──────────────────────────────────────────────────
  const newConfig = {
    meta: {
      ...(existing.meta || {}),
      version:     (existing.meta?.version || 1),
      description: existing.meta?.description || 'Конфигурация полей форм.',
      scannedAt:   new Date().toISOString().slice(0, 10),
      scannedFrom: path.basename(filePath),
    },
    groups: mergedGroups,
  };

  return newConfig;
}

// ── Запись файлов ─────────────────────────────────────────────
function writeFiles(config) {
  const json = JSON.stringify(config, null, 2);

  fs.writeFileSync(CONFIG_JSON, json, 'utf8');
  console.log(`✔ Записан: fields-config.json`);

  const js = `// СГЕНЕРИРОВАНО АВТОМАТИЧЕСКИ — не редактируйте вручную.
// Регенерировать: node scripts/scan-excel.js <путь/к/шаблону.xlsx>
/* eslint-disable */
window.FIELDS_CONFIG = ${json};
`;
  fs.writeFileSync(CONFIG_JS, js, 'utf8');
  console.log(`✔ Записан: js/fields-config.js`);
}

// ── Точка входа ───────────────────────────────────────────────
(async () => {
  console.log(`Сканирование: ${resolvedExcel}`);
  try {
    const config = await scanExcel(resolvedExcel);
    writeFiles(config);
    console.log('\nГотово. Перезапустите приложение для применения изменений.');
  } catch (err) {
    console.error('Ошибка:', err.message);
    process.exit(1);
  }
})();
