'use strict';
/**
 * datepicker.js — лёгкий кастомный календарь для полей типа "date"
 *
 * Открывается по клику на .cal-btn или на input.has-cal.
 * Формат: ДД.ММ.ГГГГ
 * Зависимостей нет. Подключается после form-builder.js.
 *
 * Архитектура:
 *  - DOM-попап строится один раз в buildPopup().
 *  - Один делегированный click-listener на popup (не на каждой кнопке).
 *  - render() обновляет только изменяемые части:
 *      календарь → текст заголовка + data-атрибуты nav + ячейки дней;
 *      YM-пикер  → классы dp-ym-active на кнопках месяцев/лет.
 */

(function () {

  // ── Русские названия ──────────────────────────────────────────
  const MONTHS_RU = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь',
  ];
  const DOW_RU = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];

  // ── Состояние ─────────────────────────────────────────────────
  let popup    = null;   // единственный DOM-попап
  let target   = null;   // текущий input
  let curYear  = 0;
  let curMonth = 0;      // 0-based
  let modeYM   = false;  // режим выбора месяца/года

  // ── Кэшированные ссылки на узлы попапа ───────────────────────
  let elCal;             // обёртка режима календаря
  let elYM;              // обёртка режима YM-пикера
  let elHeadTitle;       // <button class="dp-head-title"> — заголовок месяца/года
  let elPrev, elNext;    // кнопки навигации
  let elGrid;            // .dp-grid
  let dowHeaders;        // массив из 7 статичных <div class="dp-dow">
  let elYMMonthBtns;     // массив из 12 <button class="dp-ym-month">
  let elYMYears;         // <div class="dp-ym-years"> — перестраивается при смене диапазона
  let ymYearMin = NaN;   // нижняя граница текущего диапазона годов

  // ── Парсинг / форматирование ──────────────────────────────────
  function parseDDMMYYYY(str) {
    const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec((str || '').trim());
    if (!m) return null;
    const d = new Date(+m[3], +m[2] - 1, +m[1]);
    if (isNaN(d)) return null;
    return d;
  }
  function fmt(d) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return `${dd}.${mm}.${d.getFullYear()}`;
  }

  // ── Создаём попап один раз ────────────────────────────────────
  function buildPopup() {
    popup = document.createElement('div');
    popup.className = 'dp-popup';
    popup.setAttribute('role', 'dialog');
    popup.setAttribute('aria-modal', 'true');

    // ── Режим: обычный календарь ──────────────────────────────
    elCal = document.createElement('div');

    // Заголовок навигации
    const calHeader = document.createElement('div');
    calHeader.className = 'dp-header';

    elPrev = document.createElement('button');
    elPrev.className = 'dp-nav dp-prev';
    elPrev.title = 'Предыдущий месяц';
    elPrev.innerHTML =
      '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor"' +
      ' stroke-width="2.5" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>';

    elHeadTitle = document.createElement('button');
    elHeadTitle.className = 'dp-head-title';
    elHeadTitle.dataset.action = 'ymmode';

    elNext = document.createElement('button');
    elNext.className = 'dp-nav dp-next';
    elNext.title = 'Следующий месяц';
    elNext.innerHTML =
      '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor"' +
      ' stroke-width="2.5" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>';

    calHeader.append(elPrev, elHeadTitle, elNext);

    // Сетка (заголовки дней недели — статичные)
    elGrid = document.createElement('div');
    elGrid.className = 'dp-grid';
    dowHeaders = DOW_RU.map((label, i) => {
      const el = document.createElement('div');
      el.className = 'dp-dow' + (i >= 5 ? ' dp-weekend' : '');
      el.textContent = label;
      return el;
    });
    elGrid.append(...dowHeaders);

    // Подвал
    const calFooter = document.createElement('div');
    calFooter.className = 'dp-footer';
    const btnToday = document.createElement('button');
    btnToday.className = 'dp-today-btn';
    btnToday.textContent = 'Сегодня';
    const btnClear = document.createElement('button');
    btnClear.className = 'dp-clear-btn';
    btnClear.textContent = 'Очистить';
    calFooter.append(btnToday, btnClear);

    elCal.append(calHeader, elGrid, calFooter);

    // ── Режим: выбор месяца/года ──────────────────────────────
    elYM = document.createElement('div');
    elYM.hidden = true;

    const ymHeader = document.createElement('div');
    ymHeader.className = 'dp-header';
    const ymTitle = document.createElement('span');
    ymTitle.className = 'dp-head-title';
    ymTitle.textContent = 'Выберите месяц и год';
    ymHeader.append(ymTitle);

    const ymBody = document.createElement('div');
    ymBody.className = 'dp-ym-body';

    const elYMMonths = document.createElement('div');
    elYMMonths.className = 'dp-ym-months';
    elYMMonthBtns = MONTHS_RU.map((name, i) => {
      const btn = document.createElement('button');
      btn.className = 'dp-ym-month';
      btn.dataset.m = i;
      btn.textContent = name;
      return btn;
    });
    elYMMonths.append(...elYMMonthBtns);

    elYMYears = document.createElement('div');
    elYMYears.className = 'dp-ym-years';

    ymBody.append(elYMMonths, elYMYears);

    const ymFooter = document.createElement('div');
    ymFooter.className = 'dp-footer';
    const btnBack = document.createElement('button');
    btnBack.className = 'dp-back-btn';
    btnBack.textContent = '← Назад';
    ymFooter.append(btnBack);

    elYM.append(ymHeader, ymBody, ymFooter);

    popup.append(elCal, elYM);

    // ── Единственный делегированный обработчик ────────────────
    popup.addEventListener('click', onPopupClick);

    document.body.appendChild(popup);
  }

  function ensurePopup() {
    if (!popup) buildPopup();
  }

  // ── Делегированный обработчик кликов внутри попапа ───────────
  function onPopupClick(e) {
    const btn = e.target.closest('button');
    if (!btn) return;

    const action = btn.dataset.action;

    if (action === 'ymmode') {
      modeYM = true;
      render();
      return;
    }
    if (btn.classList.contains('dp-nav')) {
      curYear  = +btn.dataset.y;
      curMonth = +btn.dataset.m;
      render();
      return;
    }
    if (btn.classList.contains('dp-day') && !btn.classList.contains('dp-other')) {
      selectDate(new Date(+btn.dataset.y, +btn.dataset.mo, +btn.dataset.d));
      return;
    }
    if (btn.classList.contains('dp-today-btn')) {
      const now = new Date();
      curYear  = now.getFullYear();
      curMonth = now.getMonth();
      selectDate(now);
      return;
    }
    if (btn.classList.contains('dp-clear-btn')) {
      if (target) {
        target.value = '';
        target.dispatchEvent(new Event('input',  { bubbles: true }));
        target.dispatchEvent(new Event('change', { bubbles: true }));
      }
      close();
      return;
    }
    if (btn.classList.contains('dp-back-btn')) {
      modeYM = false;
      render();
      return;
    }
    if (btn.classList.contains('dp-ym-month')) {
      curMonth = +btn.dataset.m;
      modeYM   = false;
      render();
      return;
    }
    if (btn.classList.contains('dp-ym-year')) {
      curYear = +btn.dataset.y;
      modeYM  = false;
      render();
      return;
    }
  }

  // ── Позиционирование ──────────────────────────────────────────
  function positionPopup(anchor) {
    const rect = anchor.getBoundingClientRect();
    const gap  = 4;
    popup.style.left = rect.left + 'px';

    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    const ph = popup.offsetHeight || 260;

    if (spaceBelow >= ph + gap || spaceBelow >= spaceAbove) {
      popup.style.top    = (rect.bottom + gap + window.scrollY) + 'px';
      popup.style.bottom = 'auto';
    } else {
      popup.style.top    = (rect.top - ph - gap + window.scrollY) + 'px';
      popup.style.bottom = 'auto';
    }
  }

  // ── Обновление сетки дней (только изменяемая часть) ──────────
  function renderCalendar() {
    const today   = new Date();
    today.setHours(0, 0, 0, 0);
    const selected = parseDDMMYYYY(target.value);
    const selTime  = selected ? selected.getTime() : null;

    // Заголовок и кнопки навигации
    elHeadTitle.textContent = `${MONTHS_RU[curMonth]} ${curYear}`;

    const prevM = curMonth === 0  ? { y: curYear - 1, m: 11 } : { y: curYear,     m: curMonth - 1 };
    const nextM = curMonth === 11 ? { y: curYear + 1, m: 0  } : { y: curYear,     m: curMonth + 1 };
    elPrev.dataset.y = prevM.y;  elPrev.dataset.m = prevM.m;
    elNext.dataset.y = nextM.y;  elNext.dataset.m = nextM.m;

    // Вычисляем сетку
    const firstDay    = new Date(curYear, curMonth, 1);
    let   startDow    = firstDay.getDay();
    startDow = startDow === 0 ? 6 : startDow - 1; // Пн=0…Вс=6

    const daysInMonth = new Date(curYear, curMonth + 1, 0).getDate();
    const daysInPrevM = new Date(curYear, curMonth,     0).getDate();
    const totalCells  = Math.ceil((startDow + daysInMonth) / 7) * 7;
    const trailing    = totalCells - startDow - daysInMonth;

    // Строим только ячейки дней
    const cells = [];

    for (let i = 0; i < startDow; i++) {
      const btn = document.createElement('button');
      btn.className = 'dp-day dp-other';
      btn.tabIndex  = -1;
      btn.textContent = daysInPrevM - startDow + 1 + i;
      cells.push(btn);
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const dt  = new Date(curYear, curMonth, d);
      const dow = (dt.getDay() + 6) % 7;
      let cls = 'dp-day';
      if (selTime !== null && dt.getTime() === selTime) cls += ' dp-selected';
      if (dt.getTime() === today.getTime())             cls += ' dp-today';
      if (dow >= 5)                                      cls += ' dp-weekend';

      const btn = document.createElement('button');
      btn.className   = cls;
      btn.dataset.d   = d;
      btn.dataset.mo  = curMonth;
      btn.dataset.y   = curYear;
      btn.textContent = d;
      cells.push(btn);
    }

    for (let d = 1; d <= trailing; d++) {
      const btn = document.createElement('button');
      btn.className = 'dp-day dp-other';
      btn.tabIndex  = -1;
      btn.textContent = d;
      cells.push(btn);
    }

    // Заменяем только ячейки дней; заголовки дней недели остаются
    elGrid.replaceChildren(...dowHeaders, ...cells);
  }

  // ── Обновление YM-пикера (только классы) ─────────────────────
  function renderYMPicker() {
    // Активный месяц
    elYMMonthBtns.forEach((btn, i) =>
      btn.classList.toggle('dp-ym-active', i === curMonth));

    // Диапазон лет: ±7 от curYear — перестраиваем только при смене диапазона
    const minY = curYear - 7;
    if (ymYearMin !== minY) {
      ymYearMin = minY;
      const yearBtns = [];
      for (let y = minY; y <= curYear + 7; y++) {
        const btn = document.createElement('button');
        btn.className   = 'dp-ym-year';
        btn.dataset.y   = y;
        btn.textContent = y;
        yearBtns.push(btn);
      }
      elYMYears.replaceChildren(...yearBtns);
    }

    // Активный год
    elYMYears.querySelectorAll('.dp-ym-year').forEach(btn =>
      btn.classList.toggle('dp-ym-active', +btn.dataset.y === curYear));
  }

  // ── Главная функция рендера ───────────────────────────────────
  function render() {
    ensurePopup();

    if (modeYM) {
      elCal.hidden = true;
      elYM.hidden  = false;
      renderYMPicker();
    } else {
      elYM.hidden  = true;
      elCal.hidden = false;
      renderCalendar();
    }

    positionPopup(target.closest('.input-wrap') || target);
  }

  // ── Выбор даты ────────────────────────────────────────────────
  function selectDate(d) {
    if (!target) return;
    target.value = fmt(d);
    target.dispatchEvent(new Event('input',  { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
    close();
  }

  // ── Открыть ───────────────────────────────────────────────────
  function open(inputEl) {
    target = inputEl;
    modeYM = false;

    const parsed = parseDDMMYYYY(target.value);
    const base   = parsed || new Date();
    curYear  = base.getFullYear();
    curMonth = base.getMonth();

    ensurePopup();
    popup.classList.add('dp-open');
    render();
  }

  // ── Закрыть ───────────────────────────────────────────────────
  function close() {
    if (popup) popup.classList.remove('dp-open');
    target = null;
  }

  // ── Проверка: клик внутри попапа? ─────────────────────────────
  function isInsidePopup(el) {
    return popup && popup.contains(el);
  }

  // ── Делегирование кликов по документу ────────────────────────
  document.addEventListener('mousedown', function (e) {
    const calBtn   = e.target.closest('.cal-btn');
    const calInput = e.target.closest('input.has-cal');

    if (calBtn) {
      e.preventDefault();
      const inp = calBtn.closest('.input-wrap')?.querySelector('input.has-cal');
      if (!inp) return;

      if (popup && popup.classList.contains('dp-open') && target === inp) {
        close();
      } else {
        open(inp);
      }
      return;
    }

    if (calInput) {
      if (popup && popup.classList.contains('dp-open') && target === calInput) return;
      open(calInput);
      return;
    }

    if (!isInsidePopup(e.target) && popup && popup.classList.contains('dp-open')) {
      close();
    }
  }, true);

  // Escape
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && popup && popup.classList.contains('dp-open')) {
      close();
      if (target) target.focus();
    }
  });

  // ── Подстановка при ручном вводе ─────────────────────────────
  // Автовставка точек: 14 → 14. → 14.07 → 14.07. → 14.07.2026
  let _renderTimer = null;
  document.addEventListener('input', function (e) {
    const inp = e.target;
    if (!inp.classList.contains('has-cal')) return;

    let v = inp.value.replace(/[^\d.]/g, '');

    if (/^\d{2}$/.test(v))        v += '.';
    else if (/^\d{2}\.\d{2}$/.test(v)) v += '.';

    if (v !== inp.value) inp.value = v;

    // Обновляем подсветку в открытом календаре с дебаунсом
    if (popup && popup.classList.contains('dp-open') && target === inp) {
      clearTimeout(_renderTimer);
      _renderTimer = setTimeout(() => {
        const parsed = parseDDMMYYYY(v);
        if (parsed) {
          curYear  = parsed.getFullYear();
          curMonth = parsed.getMonth();
        }
        render();
      }, 150);
    }
  }, true);

})();
