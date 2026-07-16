const ExcelJS = require('exceljs');

const BLOCK_HEADERS = {
  'СДЕЛКА': 'deal',
  'ОБЪЕКТ': 'property',
  'ПРОДАВЕЦ': 'seller',
  'СОБСТВЕННИК №1': 'owner1',
  'СОБСТВЕННИК №2': 'owner2',
  'СОБСТВЕННИК №3': 'owner3',
  'ПОКУПАТЕЛЬ': 'buyer',
};

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
    if (v.result !== undefined) return String(v.result);
    if (v.text !== undefined) return String(v.text);
    if (v.richText !== undefined) {
      return v.richText.map((r) => r.text || '').join('');
    }
    return String(v);
  }
  return String(v);
}

async function readFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Файл не содержит листов');
  }

  const result = {
    deal: {},
    property: {},
    seller: {},
    owner1: {},
    owner2: {},
    owner3: {},
    buyer: {},
  };

  // rowMap: "block-fieldName" → Excel row number
  // Used by the writer to update only the correct cells in column B.
  const rowMap = {};

  // Fields computed by the app itself — skip when reading from Excel.
  const COMPUTED_FIELDS = new Set(['Комиссия агентства']);

  let currentBlock = null;

  worksheet.eachRow((row) => {
    const rawA = row.getCell(1);
    const rawB = row.getCell(2);

    const a = cellToString(rawA).trim();
    const b = cellToString(rawB).trim();

    if (!a) return;

    if (BLOCK_HEADERS[a] !== undefined) {
      currentBlock = BLOCK_HEADERS[a];
      return;
    }

    // Skip fields that are computed by the app (not read from Excel)
    if (COMPUTED_FIELDS.has(a)) return;

    if (currentBlock && result[currentBlock] !== undefined) {
      result[currentBlock][a] = b;
      rowMap[`${currentBlock}-${a}`] = row.number;
    }
  });

  return { ...result, _rowMap: rowMap };
}

module.exports = { readFile };
