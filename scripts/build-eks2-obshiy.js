/**
 * Создаёт шаблон DOCX «Договор ЭКС оказания риэлтерских услуг (2 собственника, общий)»
 * путём трансформации оригинального файла: заменяет пустые строки/___ на плейсхолдеры.
 *
 * Запуск:  node scripts/build-eks2-obshiy.js
 * Источник: attached_assets/Договор_ЭКС_2_собств_общий_1784281528395.docx
 * Результат: templates/working/Договор_ЭКС_2_собств.docx
 */

'use strict';

const fs     = require('fs');
const path   = require('path');
const PizZip = require('pizzip');

// ── Helpers ───────────────────────────────────────────────────────────────────

function getPPr(paraXml) {
  const m = paraXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : '';
}

function getFirstRPr(paraXml) {
  const m = paraXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

function replaceParaText(paraXml, newText) {
  const tagM   = paraXml.match(/^<w:p(?:\s[^>]*)?>/) ;
  const openTag = tagM ? tagM[0] : '<w:p>';
  const pPr    = getPPr(paraXml);
  const rPr    = getFirstRPr(paraXml);
  const safe   = newText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const run = `<w:r>${rPr}<w:t xml:space="preserve">${safe}</w:t></w:r>`;
  return `${openTag}${pPr}${run}</w:p>`;
}

// ── Load source DOCX ──────────────────────────────────────────────────────────

const srcPath = path.join(
  __dirname, '..', 'attached_assets',
  'Договор_ЭКС_2_собств_общий_1784281528395.docx'
);
const content = fs.readFileSync(srcPath, 'binary');
const zip     = new PizZip(content);

// Remove proofErr fragments that split placeholder text
const origXml = zip.files['word/document.xml'].asText()
  .replace(/<w:proofErr[^/]*\/>/g, '');

// ── Split into paragraphs ─────────────────────────────────────────────────────

const paraRe = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>)/g;
const paras  = [];
let m;
while ((m = paraRe.exec(origXml)) !== null) paras.push(m[1]);

// ── Replacement map  (true XML paragraph indices) ────────────────────────────
//
// Para 4 is special: the original file stores BOTH owner blanks AND the full
// ГермесГарант intro sentence in a single paragraph (with <w:br> line breaks).
// We replace the entire paragraph so both owners + intro are present.

const GERMESGARANT_INTRO =
  'в дальнейшем именуемые "Потребители", и Общество с ограниченной ответственностью ' +
  '«ГермесГарант» на основании лицензии Министерства юстиции Республики Беларусь 02240/543 ' +
  'от 12.06.2026г, договора обязательного страхования ответственности за причинение вреда ' +
  'в связи с осуществлением риэлтерской деятельности с БРУСП «Белгострах» от 26.05.2026г. ' +
  'до 25.05.2027г. (страховой полис серия БР №0005842), в лице директора Турко Ольги ' +
  'Ростиславовны, действующего на основании Устава, в дальнейшем именуемое "Исполнитель", ' +
  'совместно именуемые "Стороны", заключили настоящий договор о нижеследующем:';

const replacements = {

  // ── ШАПКА ────────────────────────────────────────────────────────────────
  0: 'ДОГОВОР №{{deal.contractNumber}}',
  3: 'г.Лида                                                                                                                                        {{deal.date}}г.',

  // Para 4: оба собственника + ГермесГарант (один абзац в оригинале)
  4: '{{seller.fullName}}, {{seller.birthDate}} г.р., паспорт {{seller.passport}}, ' +
     'выдан {{seller.passportIssuedBy}} {{seller.passportIssueDate}}, идентификационный номер ' +
     '{{seller.id}}, зарегистрированный(ая) по адресу: {{seller.address}}, ' +
     '{{owner1.fullName}}, {{owner1.birthDate}} г.р., паспорт {{owner1.passport}}, ' +
     'выдан {{owner1.passportIssuedBy}} {{owner1.passportIssueDate}}, идентификационный номер ' +
     '{{owner1.id}}, зарегистрированный(ая) по адресу: {{owner1.address}}, ' +
     GERMESGARANT_INTRO,

  // ── РАЗДЕЛ 1 ─────────────────────────────────────────────────────────────
  // 1.1 property type + address (Потребителю → Потребителям для 2 собственников)
  7:  '1.1.  Предметом настоящего договора является оказание Исполнителем Потребителям ' +
      'риэлтерских услуг по содействию в продаже принадлежащего Потребителям на праве ' +
      'собственности объекта недвижимого имущества (далее – риэлтерские услуги), ' +
      'представляющего собой – "{{property.type}}", расположенный по адресу {{property.address}}.',

  // 1.2 rooms, area, floor, material
  9:  'кол-во комнат: {{property.rooms}}; общая пл. {{property.areaTotal}} м²; ' +
      'жилая пл. {{property.areaLiving}} м²; этаж/этажность: {{property.floor}}/{{property.floors}}, материал',
  10: 'стен: {{property.wallMaterial}}, г.п. {{property.yearBuilt}}, инвентарный номер – {{property.inventoryNumber}}',

  // 1.4 price
  13: '1.4. Ориентировочная стоимость {{property.priceBYN}} ({{property.priceWords}}',
  14: ') белорусских рублей.',

  // ── РАЗДЕЛ 2 ─────────────────────────────────────────────────────────────
  // 2.2 commission
  18: '2.2. Стоимость риэлтерских услуг составляет {{commission.percent}} % от стоимости ' +
      'объекта недвижимого имущества, а именно {{commission.amountBYN}} ({{commission.amountWords}}) ' +
      'белорусских рублей, с включением в такую оплату затрат по риэлтерским услугам, указанным ' +
      'в абзацах втором — шестом пункта 1 статьи 16 Закона Республики Беларусь от 8 мая 2025 г. ' +
      '№ 71-З «О риэлтерской деятельности».',

  // ── РАЗДЕЛ 6 — агент ─────────────────────────────────────────────────────
  107: 'Ответственный риэлтер по настоящему договору {{agent.initials}}',
  109: '№ {{agent.attestationNumber}}от {{agent.attestationDate}}г, действует до {{agent.attestationExpiry}}г',
  110: 'Идентификационная пластиковая карточка {{agent.cardNumber}}',

  // ── РАЗДЕЛ 7 — реквизиты ПРОДАВЦА ────────────────────────────────────────
  113: 'ФИО: {{seller.fullName}}',
  114: '{{seller.birthDate}} г.р.',
  115: 'Регистрация: {{seller.address}}',
  116: 'Паспорт: {{seller.passport}}',
  117: 'Выдан {{seller.passportIssuedByInstrumental}} {{seller.passportIssueDate}}',
  118: 'Идентификационный номер: {{seller.id}}',
  119: 'Тел: {{seller.phone}}',
  121: 'Потребитель_______________________{{seller.initials}}',

  // ── РАЗДЕЛ 7 — реквизиты СОБСТВЕННИКА №2 ─────────────────────────────────
  132: 'ФИО: {{owner1.fullName}}',
  133: '{{owner1.birthDate}} г.р.',
  134: 'Регистрация: {{owner1.address}}',
  135: 'Паспорт: {{owner1.passport}}',
  136: 'Выдан {{owner1.passportIssuedByInstrumental}} {{owner1.passportIssueDate}}',
  137: 'Идентификационный номер: {{owner1.id}}',
  138: 'Тел: {{owner1.phone}}',
  140: 'Потребитель_______________________{{owner1.initials}}',

  // ── ПРИЛОЖЕНИЕ 1 ─────────────────────────────────────────────────────────
  147: 'к договору оказания риэлтерских услуг по содействию в продаже объектов ' +
       'недвижимого имущества №{{deal.contractNumber}} от {{deal.date}}',
  151: 'Гр. {{seller.fullName}},',
  152: '{{seller.address}}, паспорт {{seller.passport}}, идентификационный номер {{seller.id}},',

  // Appendix signature lines
  195: 'Потребитель: _______________ /{{seller.initials}}                                                                Исполнитель: _________________________О.Р.Турко',
  197: 'Потребитель: _______________ /{{owner1.initials}}',
};

// ── Apply replacements ────────────────────────────────────────────────────────

for (const [idxStr, newText] of Object.entries(replacements)) {
  const idx = parseInt(idxStr, 10);
  if (idx >= paras.length) {
    console.warn(`⚠ Para ${idx} out of range (total: ${paras.length})`);
    continue;
  }
  paras[idx] = replaceParaText(paras[idx], newText);
}

// ── Reconstruct XML ───────────────────────────────────────────────────────────

const positions = [];
const rePos = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>)/g;
let mp;
while ((mp = rePos.exec(origXml)) !== null) {
  positions.push({ start: mp.index, end: mp.index + mp[0].length });
}

const parts = [];
let lastEnd = 0;
positions.forEach((pos, i) => {
  parts.push(origXml.slice(lastEnd, pos.start));
  parts.push(paras[i]);
  lastEnd = pos.end;
});
parts.push(origXml.slice(lastEnd));
const finalXml = parts.join('');

// ── Write output ──────────────────────────────────────────────────────────────

zip.file('word/document.xml', finalXml);

const outPath = path.join(
  __dirname, '..', 'templates', 'working', 'Договор_ЭКС_2_собств.docx'
);
const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
fs.writeFileSync(outPath, buf);
console.log('✓ Шаблон создан:', outPath);
console.log('  Всего параграфов:', paras.length);
console.log('  Заменено:', Object.keys(replacements).length);
