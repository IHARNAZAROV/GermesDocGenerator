/**
 * Создаёт шаблон DOCX «Договор задатка (стандартный)» с плейсхолдерами
 * для системы ГермесГарант.
 *
 * Запуск:  node scripts/create-zadatok-template.js
 * Результат: templates/working/Договор_задатка.docx
 */

'use strict';

const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  PageOrientation,
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────

function bold(text, size = 22) {
  return new TextRun({ text, bold: true, size });
}

function run(text, opts = {}) {
  return new TextRun({ text, size: 22, ...opts });
}

function heading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 80 },
    children: [bold(text, 22)],
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    indent: { firstLine: 720 },
    ...opts,
    children,
  });
}

function paraNI(children, opts = {}) {           // no indent
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    ...opts,
    children,
  });
}

// ── document content ─────────────────────────────────────────────────────────

const doc = new Document({
  sections: [
    {
      properties: {},
      children: [

        // ── Title
        heading('ДОГОВОР ЗАДАТКА'),

        // ── City + Date line
        new Paragraph({
          spacing: { before: 80, after: 160 },
          children: [
            run('г. Лида'),
            run('\t\t\t\t\t'),
            run('{{deal.date}}'),
          ],
        }),

        // ── Preamble (parties)
        para([
          run('Мы, нижеподписавшиеся: «Задаткодатель», именуемый(ая) в дальнейшем «Покупатель», гр.\u00a0'),
          run('{{buyer.fullNameGenitive}}'),
          run(', адрес регистрации:\u00a0'),
          run('{{buyer.address}}'),
          run(', паспорт\u00a0'),
          run('{{buyer.passport}}'),
          run(', выдан\u00a0'),
          run('{{buyer.passportIssuedBy}}'),
          run(', идентификационный номер\u00a0'),
          run('{{buyer.id}}'),
          run(', и «Задаткополучатель», именуемый(ая) в дальнейшем «Продавец», гр.\u00a0'),
          run('{{seller.fullNameGenitive}}'),
          run(', зарегистрированный(ая) по адресу:\u00a0'),
          run('{{seller.address}}'),
          run(', паспорт\u00a0'),
          run('{{seller.passport}}'),
          run(', выдан\u00a0'),
          run('{{seller.passportIssuedBy}}'),
          run(', идентификационный номер\u00a0'),
          run('{{seller.id}}'),
          run(', заключили между собой настоящий договор о нижеследующем:'),
        ]),

        // ── Section 1
        heading('1. ПРЕДМЕТ ДОГОВОРА'),

        para([
          run('1.\u00a0«Продавец» получил от «Покупателя» задаток в размере:\u00a0'),
          run('{{deposit.amountBYN}}'),
          run('\u00a0('),
          run('{{deposit.amountBYNWords}}'),
          run(')\u00a0белорусских рублей 00 копеек, что эквивалентно:\u00a0'),
          run('{{deposit.amountUSD}}'),
          run('\u00a0('),
          run('{{deposit.amountUSDWords}}'),
          run(')\u00a0долларов США.'),
        ]),

        para([
          run('2.\u00a0Задаток выдан в обеспечение исполнения и в счёт причитающихся с «Покупателя» платежей по договору купли\u2013продажи\u00a0'),
          run('{{property.typeGenitive}}'),
          run('\u00a0по адресу:\u00a0'),
          run('{{property.address}}'),
          run('.'),
        ]),

        para([
          run('Стоимость объекта недвижимости составляет:\u00a0'),
          run('{{property.priceBYN}}'),
          run('\u00a0('),
          run('{{property.priceWords}}'),
          run(')\u00a0белорусских рублей 00 копеек, что эквивалентно:\u00a0'),
          run('{{property.priceUSD}}'),
          run('\u00a0('),
          run('{{property.priceWordsUSD}}'),
          run(')\u00a0долларов США.'),
        ]),

        para([
          run('1.3.\u00a0Стороны пришли к Соглашению об условиях договора, обеспечиваемого задатком.'),
        ]),

        para([
          run('1.4.\u00a0Сумма задатка передаётся «Покупателем» «Продавцу» при подписании настоящего договора.'),
        ]),

        // ── Section 2
        heading('2. ОБЯЗАННОСТИ СТОРОН'),

        para([
          run('1.\u00a0Стороны обязуются заключить договор купли-продажи, обеспеченный задатком в срок до\u00a0'),
          run('{{deal.endDate}}'),
          run('.'),
        ]),

        para([run('2.\u00a0«Продавец» гарантирует, что:')]),

        para([
          run('2.2.1.\u00a0Он сам и лица, чьё согласие необходимо для совершения сделки, и которые будут участвовать в сделке с его стороны, полностью дееспособны, действуют осознанно без принуждения и дадут необходимое согласие.'),
        ]),

        para([
          run('2.2.2.\u00a0Заявленная к продаже недвижимость до подписания настоящего договора никому другому не продана, не подарена, не заложена, не сдана в аренду, под запрещением или арестом не состоит, споров по ней не имеется, не обременена правами третьих лиц, не обладает существенными скрытыми дефектами.'),
        ]),

        para([run('2.2.3.\u00a0Все предъявленные документы являются подлинными.')]),

        // ── Section 3
        heading('3. ОТВЕТСТВЕННОСТЬ СТОРОН'),

        para([
          run('Условия возврата задатка и штрафные санкции в случае неисполнения «Продавцом» и «Покупателем» своих обязательств оговариваются:'),
        ]),

        para([
          run('3.1.\u00a0В случае неисполнения договора по вине «Покупателя» задаток не возвращается.'),
        ]),

        para([
          run('3.2.\u00a0В случае неисполнения договора по вине «Продавца» он обязан вернуть задаток в двойном размере (статья 352 п.2 ГК РБ).'),
        ]),

        para([
          run('3.3.\u00a0В случае невозможности выполнения условий договора вследствие возникновения форс\u2013мажорных обстоятельств, стороны приводятся в первоначальное положение.'),
        ]),

        para([
          run('3.4.\u00a0Договор составлен в двух экземплярах, имеющих одинаковую юридическую силу, один находится у «Покупателя», другой находится у «Продавца».'),
        ]),

        para([run('3.5.\u00a0Дополнительные условия:')]),

        para([
          run('3.5.1.\u00a0В квартире остаётся следующая мебель:\u00a0'),
          run('{{deal.furniture}}'),
          run('.'),
        ]),

        para([
          run('3.5.2.\u00a0Оплата услуг БТИ осуществляется\u00a0'),
          run('{{deal.btiPayment}}'),
          run('.'),
        ]),

        // ── Section 4 — Signatures (table)
        heading('4. ПОДПИСИ СТОРОН'),

        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top:           { style: BorderStyle.SINGLE, size: 4 },
            bottom:        { style: BorderStyle.SINGLE, size: 4 },
            left:          { style: BorderStyle.SINGLE, size: 4 },
            right:         { style: BorderStyle.SINGLE, size: 4 },
            insideH:       { style: BorderStyle.SINGLE, size: 4 },
            insideV:       { style: BorderStyle.SINGLE, size: 4 },
          },
          rows: [
            // Header row
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [bold('«Продавец»')],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [run('(задаткополучатель)')],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 50, type: WidthType.PERCENTAGE },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [bold('«Покупатель»')],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [run('(задаткодатель)')],
                    }),
                  ],
                }),
              ],
            }),
            // Name row
            new TableRow({
              children: [
                new TableCell({
                  children: [paraNI([run('гр. '), run('{{seller.fullName}}')])],
                }),
                new TableCell({
                  children: [paraNI([run('гр. '), run('{{buyer.fullName}}')])],
                }),
              ],
            }),
            // Address row
            new TableRow({
              children: [
                new TableCell({
                  children: [paraNI([run('{{seller.address}}')])],
                }),
                new TableCell({
                  children: [paraNI([run('{{buyer.address}}')])],
                }),
              ],
            }),
            // Passport row
            new TableRow({
              children: [
                new TableCell({
                  children: [paraNI([run('{{seller.passport}}')])],
                }),
                new TableCell({
                  children: [paraNI([run('{{buyer.passport}}')])],
                }),
              ],
            }),
            // ID row
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    paraNI([run('идентификационный номер')]),
                    paraNI([run('{{seller.id}}')]),
                  ],
                }),
                new TableCell({
                  children: [
                    paraNI([run('идентификационный номер')]),
                    paraNI([run('{{buyer.id}}')]),
                  ],
                }),
              ],
            }),
            // Signature row
            new TableRow({
              children: [
                new TableCell({
                  children: [paraNI([run('Подпись _________________')])],
                }),
                new TableCell({
                  children: [paraNI([run('Подпись _________________')])],
                }),
              ],
            }),
          ],
        }),

      ],
    },
  ],
});

// ── Write output ─────────────────────────────────────────────────────────────

const outPath = path.join(__dirname, '..', 'templates', 'working', 'Договор_задатка.docx');

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outPath, buffer);
  console.log('✓ Шаблон создан:', outPath);
}).catch((err) => {
  console.error('✗ Ошибка:', err.message);
  process.exit(1);
});
