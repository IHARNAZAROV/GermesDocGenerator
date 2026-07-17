/**
 * Создаёт шаблон DOCX «Договор оказания риэлтерских услуг (2 собственника, общий)»
 * путём трансформации оригинального файла: заменяет пустые строки/___ на плейсхолдеры.
 *
 * Запуск:  node scripts/build-dkp2-obshiy.js
 * Источник: attached_assets/Договор_2_собств_общий_1784277589269.docx
 * Результат: templates/working/Договор_2_собств_общий.docx
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const PizZip = require('pizzip');

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Extract pPr block (paragraph properties) from a paragraph XML string */
function getPPr(paraXml) {
  const m = paraXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : '';
}

/** Extract rPr block (run properties) from the first run in a paragraph */
function getFirstRPr(paraXml) {
  const m = paraXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

/** Build a replacement paragraph: keep pPr + rPr, replace text with newText */
function makePara(paraXml, newText) {
  const pPr  = getPPr(paraXml);
  const rPr  = getFirstRPr(paraXml);
  // XML-escape the text ({{ and }} are fine — only <>&"' need escaping)
  const safe = newText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Split on spaces that appear after long runs to avoid Word wrapping issues,
  // but simplest: emit a single run.
  const run = `<w:r>${rPr}<w:t xml:space="preserve">${safe}</w:t></w:r>`;
  return `<w:p>${pPr}${run}</w:p>`;
}

/** Replace all paragraph text while keeping its XML tag + attributes intact */
function replaceParaText(paraXml, newText) {
  // Strip the outer <w:p ...> tag for analysis, rebuild below
  const tagM = paraXml.match(/^<w:p(?:\s[^>]*)?>/) ;
  const openTag = tagM ? tagM[0] : '<w:p>';

  const pPr  = getPPr(paraXml);
  const rPr  = getFirstRPr(paraXml);
  const safe = newText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const run = `<w:r>${rPr}<w:t xml:space="preserve">${safe}</w:t></w:r>`;
  return `${openTag}${pPr}${run}</w:p>`;
}

// ── Load source DOCX ─────────────────────────────────────────────────────────

const srcPath = path.join(__dirname, '..', 'attached_assets', 'Договор_2_собств_общий_1784277589269.docx');
const content  = fs.readFileSync(srcPath, 'binary');
const zip      = new PizZip(content);

// Remove proofErr elements that can fragment placeholders
let xml = zip.files['word/document.xml'].asText();
xml = xml.replace(/<w:proofErr[^/]*\/>/g, '');

// ── Split into paragraphs (keep ALL para XML for reconstruction) ──────────────

const paraRe = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>)/g;
const paras  = [];       // array of paragraph XML strings (all, including empty)
let m;
while ((m = paraRe.exec(xml)) !== null) paras.push(m[1]);

// Helper: extract readable text from a paragraph XML (for logging/debugging)
function paraText(p) {
  const noPPr = p.replace(/<w:pPr[\s\S]*?<\/w:pPr>/g, '');
  let t = '', r;
  const re = /<w:t(?=>|[ ])[^>]*>([\s\S]*?)<\/w:t>/g;
  while ((r = re.exec(noPPr)) !== null) t += r[1];
  return t.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
}

// Build full text array for index-based lookups
const texts = paras.map(paraText);

// ── Index-based replacements ──────────────────────────────────────────────────
// (verified against extracted paragraph indices above)

const replacements = {

  // ── ШАПКА ────────────────────────────────────────────────────────
  0: 'ДОГОВОР №{{deal.contractNumber}}',

  // Date line
  3: 'г.Лида                                                                                                                                        {{deal.date}}г.',

  // Seller (Потребитель 1) preamble line
  4: '{{seller.fullName}}, {{seller.birthDate}} г.р., паспорт {{seller.passport}}, выдан {{seller.passportIssuedBy}} {{seller.passportIssueDate}}, идентификационный номер {{seller.id}}, зарегистрированный(ая) по адресу: {{seller.address}},',

  // Owner 1 (Потребитель 2) preamble line
  5: '{{owner1.fullName}}, {{owner1.birthDate}} г.р., паспорт {{owner1.passport}}, выдан {{owner1.passportIssuedBy}} {{owner1.passportIssueDate}}, идентификационный номер {{owner1.id}}, зарегистрированный(ая) по адресу: {{owner1.address}},',

  // ── РАЗДЕЛ 1 ─────────────────────────────────────────────────────
  // 1.1 property type + address (idx 9, not 8 — idx 8 is "1. ПРЕДМЕТ ДОГОВОРА")
  9:  '1.1.  Предметом настоящего договора является оказание Исполнителем Потребителям риэлтерских услуг по содействию в продаже принадлежащего Потребителям на праве собственности объекта недвижимого имущества (далее – риэлтерские услуги), представляющего собой – "{{property.type}}", расположенный по адресу {{property.address}}.',

  // 1.2 rooms, area, floor, material
  11: 'кол-во комнат: {{property.rooms}}; общая пл. {{property.areaTotal}} м²; жилая {{property.areaLiving}} м²; этаж/этажность: {{property.floor}}/{{property.floors}}, материал',
  12: 'стен: {{property.wallMaterial}}, г.п. {{property.yearBuilt}}, инвентарный номер – {{property.inventoryNumber}}',

  // 1.4 price
  15: '1.4. Ориентировочная стоимость {{property.priceBYN}} ({{property.priceWords}}',
  16: ') белорусских рублей.',

  // ── РАЗДЕЛ 2 ─────────────────────────────────────────────────────
  // 2.2 commission
  20: '2.2. Стоимость риэлтерских услуг составляет {{commission.percent}} % от стоимости объекта недвижимого имущества, а именно {{commission.amountBYN}} ({{commission.amountWords}}) белорусских рублей, с включением в такую оплату затрат по риэлтерским услугам, указанным в абзацах втором — шестом пункта 1 статьи 16 Закона Республики Беларусь от 8 мая 2025 г. № 71-З «О риэлтерской деятельности».',

  // ── РАЗДЕЛ 6 — агент ─────────────────────────────────────────────
  110: 'Ответственный риэлтер по настоящему договору {{agent.initials}}',
  112: '№ {{agent.attestationNumber}}от {{agent.attestationDate}}г, действует до {{agent.attestationExpiry}}г',
  113: 'Идентификационная пластиковая карточка {{agent.cardNumber}}',

  // ── РАЗДЕЛ 7 — реквизиты ПРОДАВЦА ───────────────────────────────
  126: 'ФИО: {{seller.fullName}}',
  127: '{{seller.birthDate}} г.р.',
  128: 'Регистрация: {{seller.address}}',
  129: 'Паспорт: {{seller.passport}}',
  130: 'Выдан {{seller.passportIssuedByInstrumental}} {{seller.passportIssueDate}}',
  131: 'Идентификационный номер: {{seller.id}}',
  132: 'Тел: {{seller.phone}}',
  133: 'Потребитель_______________________{{seller.initials}}',

  // ── РАЗДЕЛ 7 — реквизиты СОБСТВЕННИКА №2 ────────────────────────
  143: 'ФИО: {{owner1.fullName}}',
  144: '{{owner1.birthDate}} г.р.',
  145: 'Регистрация: {{owner1.address}}',
  146: 'Паспорт: {{owner1.passport}}',
  147: 'Выдан {{owner1.passportIssuedByInstrumental}} {{owner1.passportIssueDate}}',
  148: 'Идентификационный номер: {{owner1.id}}',
  149: 'Тел: {{owner1.phone}}',
  150: 'Потребитель_______________________{{owner1.initials}}',

  // ── ПРИЛОЖЕНИЕ 1 ─────────────────────────────────────────────────
  196: 'к договору оказания риэлтерских услуг по содействию в продаже объектов недвижимого имущества №{{deal.contractNumber}} от {{deal.date}}',
  200: 'Гр. {{seller.fullName}},',
  202: '{{seller.address}}, паспорт {{seller.passport}}, идентификационный номер {{seller.id}},',

  // Appendix signature lines
  246: 'Потребитель: _______________ /{{seller.initials}}                                                                Исполнитель: _________________________О.Р.Турко',
  248: 'Потребитель: _______________ /{{owner1.initials}}',
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

// Re-stitch everything outside paragraphs
// Strategy: replace all original paragraphs back into the xml string
let rebuilt = xml;
let offset  = 0;
// We need to collect all match positions from the original xml
const rePos = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>)/g;
const positions = [];
let mp;
// Re-run on original xml to get positions
const origXml = zip.files['word/document.xml'].asText().replace(/<w:proofErr[^/]*\/>/g, '');
const rePos2 = /(<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>)/g;
while ((mp = rePos2.exec(origXml)) !== null) {
  positions.push({ start: mp.index, end: mp.index + mp[0].length });
}

// Build the final XML by replacing each paragraph position
let parts = [];
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

const outPath = path.join(__dirname, '..', 'templates', 'working', 'Договор_2_собств_общий.docx');
const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
fs.writeFileSync(outPath, buf);
console.log('✓ Шаблон создан:', outPath);
console.log('  Всего параграфов:', paras.length);
console.log('  Заменено:', Object.keys(replacements).length);
