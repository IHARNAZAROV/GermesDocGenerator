'use strict';

// ============================================================
//  Вычисление комиссии агентства по тарифной таблице
//  Алгоритм:
//    1. priceBYN / baseValue → количество базовых величин
//    2. Найти строку в таблице brackets по этому количеству
//    3. commission = priceBYN * percent / 100
//  Источник: config/commission-rates.json — единственный источник тарифов.
//  Хардкод убран. При изменении JSON данные подтянутся автоматически.
// ============================================================

// Временный конфиг до загрузки JSON (не должен использоваться в расчётах,
// но защищает от краша если calculateCommission вызван до ready-события)
window.COMMISSION_CONFIG = {
  baseValue: 45,
  brackets: [],
  _loaded: false,
};

// Промис, который резолвится когда конфиг загружен
window.COMMISSION_CONFIG_READY = (async () => {
  try {
    const response = await fetch('./config/commission-rates.json');
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    window.COMMISSION_CONFIG = {
      baseValue: data.baseValue,
      brackets:  data.brackets,
      _loaded:   true,
    };
  } catch (err) {
    console.error('[commission] Не удалось загрузить commission-rates.json:', err);
    // Оставляем временный конфиг с хардкодом только как крайний fallback
    window.COMMISSION_CONFIG = {
      baseValue: 45,
      brackets: [
        { upTo:   700, percent: 6.0 },
        { upTo:  2000, percent: 5.0 },
        { upTo:  3000, percent: 4.0 },
        { upTo:  4200, percent: 3.0 },
        { upTo:  5000, percent: 2.5 },
        { upTo:  5800, percent: 2.4 },
        { upTo:  6600, percent: 2.3 },
        { upTo:  7500, percent: 2.2 },
        { upTo:  8300, percent: 2.1 },
        { upTo:  9100, percent: 2.0 },
        { upTo: 10000, percent: 1.9 },
        { upTo: 10500, percent: 1.8 },
        { upTo: 11600, percent: 1.7 },
        { upTo: 12400, percent: 1.6 },
        { upTo:  null, percent: 1.5 },
      ],
      _loaded: false,
      _fallback: true,
    };
  }
})();

/**
 * Вычисляет комиссию агентства.
 * @param {number} priceBYN   — стоимость объекта в BYN
 * @param {number} [baseValue] — базовая величина (по умолчанию из COMMISSION_CONFIG)
 * @returns {{ percent: number, amountBYN: string, amountBYNRaw: number, amountWords: string, baseUnits: number }}
 */
window.calculateCommission = function calculateCommission(priceBYN, baseValue) {
  const bv = baseValue || window.COMMISSION_CONFIG.baseValue;

  if (!priceBYN || isNaN(priceBYN) || priceBYN <= 0) {
    return { percent: 0, amountBYN: '', amountBYNRaw: 0, amountWords: '', baseUnits: 0 };
  }

  const baseUnits = priceBYN / bv;

  // Найти нужную строку тарифной таблицы
  const bracket = window.COMMISSION_CONFIG.brackets.find(
    (b) => b.upTo === null || baseUnits <= b.upTo
  );
  const percent = bracket ? bracket.percent : 1.5;

  const amountBYNRaw = Math.round(priceBYN * percent) / 100;
  // Форматируем до 2 знаков после запятой, убираем хвостовые нули
  const amountBYN = amountBYNRaw % 1 === 0
    ? String(amountBYNRaw)
    : amountBYNRaw.toFixed(2).replace(/\.?0+$/, '');

  // Прописью — используем существующую функцию moneyToText, если доступна
  const amountWords = (typeof window.moneyToText === 'function' && amountBYNRaw > 0)
    ? window.moneyToText(String(amountBYNRaw))
    : '';

  return { percent, amountBYN, amountBYNRaw, amountWords, baseUnits };
};
