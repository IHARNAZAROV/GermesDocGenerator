'use strict';
/**
 * excel-scanner.js — ядро сканера шаблона.
 * Используется как из scripts/scan-excel.js (CLI),
 * так и из main.js (IPC-обработчик кнопки «Обновить шаблон»).
 */

const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');

const ROOT        = path.resolve(__dirname, '..');
const CONFIG_JSON = path.join(ROOT, 'fields-config.json');
const CONFIG_JS   = path.join(ROOT, 'js', 'fields-config.js');

// ── Утилиты ───────────────────────────────────────────────────
function cellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) return v.toLocaleDateString('ru-RU');
    if (v.result   !== undefined) return String(v.result);
    if (v.text     !== undefined) return String(v.text);
    if (v.richText !== undefined) return v.richText.map(r => r.text || '').join('');
    return String(v);
  }
  return String(v);
}

function inferType(key) {
  return key.toLowerCase().startsWith('дата') ? 'date' : 'text';
}

// ── Главная функция ───────────────────────────────────────────
/**
 * Сканирует Excel, обновляет fields-config.json и js/fields-config.js.
 * @param {string} excelPath  абсолютный путь к Excel-файлу
 * @returns {{ added: string[], removed: string[], total: number, scannedFrom: string }}
 */
async function scanAndUpdate(excelPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(excelPath);

  // Загружаем существующий конфиг
  let existing = { meta: {}, groups: [] };
  if (fs.existsSync(CONFIG_JSON)) {
    try { existing = JSON.parse(fs.readFileSync(CONFIG_JSON, 'utf8')); }
    catch (_) { /* начинаем заново */ }
  }

  // Карта excelHeader → groupId из конфига
  const headerToGroupId = {};
  for (const g of (existing.groups || [])) {
    if (g.excelHeader) headerToGroupId[g.excelHeader] = g.id;
  }

  // Индекс существующих полей: groupId → { key → fieldObj }
  const existingGroups = {};
  for (const g of (existing.groups || [])) {
    existingGroups[g.id] = {};
    for (const f of (g.fields || [])) existingGroups[g.id][f.key] = f;
  }

  // Сканируем листы
  const scannedGroups = {}; // groupId → { excelHeader, fields: [{ key }] }
  wb.eachSheet((ws) => {
    let curGroupId = null;
    ws.eachRow((row) => {
      const a = cellText(row.getCell(1)).trim();
      if (!a) return;
      const gid = headerToGroupId[a];
      if (gid !== undefined) {
        curGroupId = gid;
        if (!scannedGroups[curGroupId]) scannedGroups[curGroupId] = { excelHeader: a, fields: [] };
        return;
      }
      if (curGroupId) scannedGroups[curGroupId].fields.push({ key: a });
    });
  });

  // Слияние
  const added   = [];
  const removed = [];
  const mergedGroups = [];

  for (const existingGroup of (existing.groups || [])) {
    const gid     = existingGroup.id;
    const scanned = scannedGroups[gid];
    const oldByKey = existingGroups[gid] || {};

    // Если группа исчезла и не содержит computed-полей — пропускаем
    if (!scanned && Object.values(oldByKey).every(f => !f.computed)) {
      Object.keys(oldByKey).forEach(k => removed.push(`${gid}.${k}`));
      continue;
    }

    const scannedKeys  = new Set((scanned?.fields || []).map(f => f.key));
    const mergedFields = [];

    // Сохраняем существующие поля
    for (const oldField of (existingGroup.fields || [])) {
      if (oldField.computed) {
        mergedFields.push(oldField);
        continue;
      }
      if (scannedKeys.has(oldField.key)) {
        mergedFields.push(oldField);
        scannedKeys.delete(oldField.key);
      } else {
        removed.push(`${gid}.${oldField.key}`);
      }
    }

    // Добавляем новые поля из Excel
    if (scanned) {
      for (const { key } of scanned.fields) {
        if (!scannedKeys.has(key)) continue;
        added.push(`${gid}.${key}`);
        mergedFields.push({ key, label: key + ':', type: inferType(key) });
      }
    }

    mergedGroups.push({ ...existingGroup, fields: mergedFields });
  }

  const newConfig = {
    meta: {
      ...(existing.meta || {}),
      scannedAt:   new Date().toISOString().slice(0, 10),
      scannedFrom: path.basename(excelPath),
    },
    groups: mergedGroups,
  };

  // Записываем файлы
  const json = JSON.stringify(newConfig, null, 2);
  fs.writeFileSync(CONFIG_JSON, json, 'utf8');
  fs.writeFileSync(
    CONFIG_JS,
    '// СГЕНЕРИРОВАНО АВТОМАТИЧЕСКИ — не редактируйте вручную.\n'
    + '// Регенерировать: node scripts/scan-excel.js <путь/к/шаблону.xlsx>\n'
    + '/* eslint-disable */\n'
    + `window.FIELDS_CONFIG = ${json};\n`,
    'utf8'
  );

  const total = mergedGroups.reduce((s, g) => s + g.fields.length, 0);
  return { added, removed, total, scannedFrom: path.basename(excelPath) };
}

module.exports = { scanAndUpdate };
