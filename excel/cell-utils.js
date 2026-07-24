'use strict';

/**
 * Общая утилита для ExcelJS: преобразует значение ячейки в строку.
 * Обрабатывает Date, richText, formula result, plain value.
 * Единственная точка исправления при изменении формата Excel.
 */
function cellToString(cell) {
  if (cell === null || cell === undefined) return '';
  const v = cell.value !== undefined ? cell.value : cell;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) {
      const d = v.getDate().toString().padStart(2, '0');
      const m = (v.getMonth() + 1).toString().padStart(2, '0');
      const y = v.getFullYear();
      return `${d}.${m}.${y}`;
    }
    if (v.result   !== undefined) return String(v.result);
    if (v.text     !== undefined) return String(v.text);
    if (v.richText !== undefined) return v.richText.map(r => r.text || '').join('');
    return String(v);
  }
  return String(v);
}

module.exports = { cellToString };
