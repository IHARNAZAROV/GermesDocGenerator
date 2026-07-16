'use strict';

/**
 * money-to-text.js
 * Converts a BYN (Belarusian ruble) amount to a Russian-language written form.
 * Exposed as: window.moneyToText(amount) → string
 *
 * Supports: 0.00 … 999 999 999.99 BYN
 * No external dependencies.
 */
(function () {

  // ── vocabulary ──────────────────────────────────────────────
  const ONES_M = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'];
  const ONES_F = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'];
  const TEENS  = [
    'десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать',
    'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать',
  ];
  const TENS   = ['', 'десять', 'двадцать', 'тридцать', 'сорок', 'пятьдесят',
                  'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
  const HUNDS  = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот',
                  'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];

  // ── helpers ─────────────────────────────────────────────────

  /**
   * Returns the correct Russian plural form for n.
   * @param {number} n
   * @param {string} f1  form for 1      (e.g. "рубль")
   * @param {string} f2  form for 2–4    (e.g. "рубля")
   * @param {string} f5  form for 5–20   (e.g. "рублей")
   */
  function pluralForm(n, f1, f2, f5) {
    const mod100 = Math.abs(n) % 100;
    const mod10  = Math.abs(n) % 10;
    if (mod100 >= 11 && mod100 <= 19) return f5;
    if (mod10 === 1) return f1;
    if (mod10 >= 2 && mod10 <= 4) return f2;
    return f5;
  }

  /**
   * Converts 0–999 to Russian words.
   * @param {number} n       Integer 0–999
   * @param {'m'|'f'} gender Grammatical gender for 1 and 2
   */
  function threeDigits(n, gender) {
    if (n === 0) return '';
    const parts = [];
    const h = Math.floor(n / 100);
    const rem = n % 100;
    if (h) parts.push(HUNDS[h]);
    if (rem >= 10 && rem <= 19) {
      parts.push(TEENS[rem - 10]);
    } else {
      const t = Math.floor(rem / 10);
      const o = rem % 10;
      if (t) parts.push(TENS[t]);
      if (o) parts.push(gender === 'f' ? ONES_F[o] : ONES_M[o]);
    }
    return parts.join(' ');
  }

  /**
   * Converts the integer ruble amount to lowercase Russian words
   * (without the currency label).
   */
  function rublesInWords(rubles) {
    if (rubles === 0) return 'ноль';

    const billions  = Math.floor(rubles / 1_000_000_000);
    const millions  = Math.floor((rubles % 1_000_000_000) / 1_000_000);
    const thousands = Math.floor((rubles % 1_000_000) / 1_000);
    const remainder = rubles % 1_000;

    const parts = [];

    if (billions) {
      const w = threeDigits(billions, 'm');
      parts.push(w + ' ' + pluralForm(billions, 'миллиард', 'миллиарда', 'миллиардов'));
    }
    if (millions) {
      const w = threeDigits(millions, 'm');
      parts.push(w + ' ' + pluralForm(millions, 'миллион', 'миллиона', 'миллионов'));
    }
    if (thousands) {
      const w = threeDigits(thousands, 'f');           // тысяча — feminine
      parts.push(w + ' ' + pluralForm(thousands, 'тысяча', 'тысячи', 'тысяч'));
    }
    if (remainder) {
      parts.push(threeDigits(remainder, 'm'));          // рубль — masculine
    }

    return parts.join(' ');
  }

  // ── main export ─────────────────────────────────────────────

  /**
   * Converts an amount to a Russian written form for BYN.
   * @param {number|string} amount  e.g. 105000.35 or "105000,35"
   * @returns {string}  e.g. "Сто пять тысяч белорусских рублей 35 копеек"
   *                    Returns '' for invalid input.
   */
  function moneyToText(amount) {
    // Normalise: accept comma as decimal separator
    const str = String(amount).replace(',', '.').trim();
    const num = parseFloat(str);
    if (isNaN(num) || num < 0) return '';

    // Round to 2 decimal places to avoid floating-point drift
    const totalKopecks = Math.round(num * 100);
    const rubles  = Math.floor(totalKopecks / 100);
    const kopecks = totalKopecks % 100;

    const rubWord = pluralForm(rubles,
      'белорусский рубль',
      'белорусских рубля',
      'белорусских рублей',
    );

    const kopStr  = String(kopecks).padStart(2, '0');
    const kopWord = pluralForm(kopecks, 'копейка', 'копейки', 'копеек');

    const raw = rublesInWords(rubles) + ' ' + rubWord + ' ' + kopStr + ' ' + kopWord;

    // Capitalise the very first letter of the sentence
    return raw.charAt(0).toUpperCase() + raw.slice(1);
  }

  window.moneyToText = moneyToText;

  // ── USD variant ─────────────────────────────────────────────
  /**
   * Converts an amount to a Russian written form for USD.
   * @param {number|string} amount  e.g. 105000 or "105000.50"
   * @returns {string}  e.g. "Сто пять тысяч долларов США 50 центов"
   *                    Returns '' for invalid input.
   */
  function moneyToTextUSD(amount) {
    const str = String(amount).replace(',', '.').trim();
    const num = parseFloat(str);
    if (isNaN(num) || num < 0) return '';

    const dollars = Math.floor(num);

    const dollarWord = pluralForm(dollars, 'доллар США', 'доллара США', 'долларов США');

    const dollarsText = dollars === 0 ? 'ноль' : rublesInWords(dollars);
    const raw = dollarsText + ' ' + dollarWord;

    return raw.charAt(0).toUpperCase() + raw.slice(1);
  }

  window.moneyToTextUSD = moneyToTextUSD;

})();
