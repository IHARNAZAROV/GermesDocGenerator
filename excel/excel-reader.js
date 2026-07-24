const ExcelJS = require('exceljs');
const { cellToString } = require('./cell-utils');

const BLOCK_HEADERS = {
  'СДЕЛКА': 'deal',
  'ОБЪЕКТ': 'property',
  'ПРОДАВЕЦ': 'seller',
  'СОБСТВЕННИК №1': 'owner1',
  'СОБСТВЕННИК №2': 'owner2',
  'СОБСТВЕННИК №3': 'owner3',
  'ПОКУПАТЕЛЬ': 'buyer',
};


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

module.exports = { readFile, BLOCK_HEADERS };
