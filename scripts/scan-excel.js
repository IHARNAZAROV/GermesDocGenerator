#!/usr/bin/env node
/**
 * scan-excel.js — CLI-обёртка над excel/excel-scanner.js
 *
 * Использование:
 *   node scripts/scan-excel.js <путь/к/шаблону.xlsx>
 *   npm run scan <путь/к/шаблону.xlsx>
 */

'use strict';

const path    = require('path');
const fs      = require('fs');
const { scanAndUpdate } = require('../excel/excel-scanner');

const excelPath = process.argv[2];
if (!excelPath) {
  console.error('Использование: node scripts/scan-excel.js <путь/к/шаблону.xlsx>');
  process.exit(1);
}

const resolved = path.resolve(excelPath);
if (!fs.existsSync(resolved)) {
  console.error(`Файл не найден: ${resolved}`);
  process.exit(1);
}

(async () => {
  console.log(`Сканирование: ${resolved}`);
  try {
    const { added, removed, total, scannedFrom } = await scanAndUpdate(resolved);
    if (added.length)   console.log('  Добавлено:', added.join(', '));
    if (removed.length) console.log('  Удалено:  ', removed.join(', '));
    console.log(`✔ Шаблон: ${scannedFrom} | Полей всего: ${total}`);
    console.log('✔ Записан: fields-config.json');
    console.log('✔ Записан: js/fields-config.js');
    console.log('\nГотово. Перезапустите приложение для применения изменений.');
  } catch (err) {
    console.error('Ошибка:', err.message);
    process.exit(1);
  }
})();
