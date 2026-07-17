'use strict';
/**
 * datepicker.js — лёгкий кастомный календарь для полей типа "date"
 *
 * Открывается по клику на .cal-btn или на input.has-cal.
 * Формат: ДД.ММ.ГГГГ
 * Зависимостей нет. Подключается после form-builder.js.
 */

(function () {

  // ── Русские названия ──────────────────────────────────────────
  const MONTHS_RU = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'
  ];
  const DOW_RU = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];

  // ── Состояние ─────────────────────────────────────────────────
  let popup    = null;   // единственный DOM-попап
  let target   = null;   // текущий input
  let curYear  = 0;
  let curMonth = 0;      // 0-based
  let modeYM   = false;  // режим выбора месяца/года

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
  function ensurePopup() {
    if (popup) return;
    popup = document.createElement('div');
    popup.className = 'dp-popup';
    popup.setAttribute('role', 'dialog');
    popup.setAttribute('aria-modal', 'true');
    document.body.appendChild(popup);
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

  // ── Рендер режима выбора года/месяца ─────────────────────────
  function renderYMPicker() {
    const selected = parseDDMMYYYY(target.value);
    const selY = selected ? selected.getFullYear() : null;
    const selM = selected ? selected.getMonth()    : null;

    // Диапазон лет: ±10 от текущего
    const baseYear = curYear;
    const years = [];
    for (let y = baseYear - 7; y <= baseYear + 7; y++) years.push(y);

    let html = `<div class="dp-header">
      <span class="dp-head-title">Выберите месяц и год</span>
    </div>
    <div class="dp-ym-body">
      <div class="dp-ym-months">`;

    MONTHS_RU.forEach((name, i) => {
      const active = (i === curMonth) ? ' dp-ym-active' : '';
      html += `<button class="dp-ym-month${active}" data-m="${i}">${name}</button>`;
    });

    html += `</div><div class="dp-ym-years">`;

    years.forEach(y => {
      const active = (y === curYear) ? ' dp-ym-active' : '';
      html += `<button class="dp-ym-year${active}" data-y="${y}">${y}</button>`;
    });

    html += `</div></div>
    <div class="dp-footer">
      <button class="dp-back-btn">← Назад</button>
    </div>`;

    popup.innerHTML = html;

    popup.querySelector('.dp-back-btn').addEventListener('click', () => {
      modeYM = false;
      render();
    });
    popup.querySelectorAll('.dp-ym-month').forEach(btn => {
      btn.addEventListener('click', () => {
        curMonth = +btn.dataset.m;
        modeYM = false;
        render();
      });
    });
    popup.querySelectorAll('.dp-ym-year').forEach(btn => {
      btn.addEventListener('click', () => {
        curYear = +btn.dataset.y;
        // scroll to keep year visible and rerender
        modeYM = false;
        render();
      });
    });
  }

  // ── Рендер обычного календаря ─────────────────────────────────
  function render() {
    ensurePopup();
    if (modeYM) { renderYMPicker(); positionPopup(target.closest('.input-wrap') || target); return; }

    const today    = new Date();
    today.setHours(0,0,0,0);
    const selected = parseDDMMYYYY(target.value);
    const selTime  = selected ? selected.getTime() : null;

    // Первый день месяца → определяем с какого дня недели начинать
    const firstDay = new Date(curYear, curMonth, 1);
    // В России неделя начинается с Пн (1). JS: 0=Вс, 1=Пн…
    let startDow = firstDay.getDay(); // 0-6
    startDow = (startDow === 0) ? 6 : startDow - 1; // 0=Пн…6=Вс

    const daysInMonth  = new Date(curYear, curMonth + 1, 0).getDate();
    const daysInPrevM  = new Date(curYear, curMonth, 0).getDate();

    // Навигация
    const prevMonth = curMonth === 0  ? { y: curYear - 1, m: 11 } : { y: curYear, m: curMonth - 1 };
    const nextMonth = curMonth === 11 ? { y: curYear + 1, m: 0  } : { y: curYear, m: curMonth + 1 };

    let html = `<div class="dp-header">
      <button class="dp-nav dp-prev" data-y="${prevMonth.y}" data-m="${prevMonth.m}" title="Предыдущий месяц">
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <button class="dp-head-title" data-action="ymmode">${MONTHS_RU[curMonth]} ${curYear}</button>
      <button class="dp-nav dp-next" data-y="${nextMonth.y}" data-m="${nextMonth.m}" title="Следующий месяц">
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>
    <div class="dp-grid">`;

    // Заголовок дней недели
    DOW_RU.forEach((d, i) => {
      const we = i >= 5 ? ' dp-weekend' : '';
      html += `<div class="dp-dow${we}">${d}</div>`;
    });

    // Дни предыдущего месяца (серые)
    for (let i = 0; i < startDow; i++) {
      const d = daysInPrevM - startDow + 1 + i;
      html += `<button class="dp-day dp-other" tabindex="-1">${d}</button>`;
    }

    // Дни текущего месяца
    for (let d = 1; d <= daysInMonth; d++) {
      const dt = new Date(curYear, curMonth, d);
      const dow = (dt.getDay() + 6) % 7; // 0=Пн…6=Вс
      const isToday = dt.getTime() === today.getTime();
      const isSel   = selTime !== null && dt.getTime() === selTime;
      const isWE    = dow >= 5;

      let cls = 'dp-day';
      if (isSel)   cls += ' dp-selected';
      if (isToday) cls += ' dp-today';
      if (isWE)    cls += ' dp-weekend';

      html += `<button class="${cls}" data-d="${d}" data-mo="${curMonth}" data-y="${curYear}">${d}</button>`;
    }

    // Дни следующего месяца (серые)
    const totalCells = Math.ceil((startDow + daysInMonth) / 7) * 7;
    const trailingDays = totalCells - startDow - daysInMonth;
    for (let d = 1; d <= trailingDays; d++) {
      html += `<button class="dp-day dp-other" tabindex="-1">${d}</button>`;
    }

    html += `</div>
    <div class="dp-footer">
      <button class="dp-today-btn">Сегодня</button>
      <button class="dp-clear-btn">Очистить</button>
    </div>`;

    popup.innerHTML = html;

    // ── Обработчики ───────────────────────────────────────────
    popup.querySelector('[data-action="ymmode"]').addEventListener('click', () => {
      modeYM = true;
      render();
    });

    popup.querySelectorAll('.dp-nav').forEach(btn => {
      btn.addEventListener('click', () => {
        curYear  = +btn.dataset.y;
        curMonth = +btn.dataset.m;
        render();
      });
    });

    popup.querySelectorAll('.dp-day:not(.dp-other)').forEach(btn => {
      btn.addEventListener('click', () => {
        const d = +btn.dataset.d;
        const m = +btn.dataset.mo;
        const y = +btn.dataset.y;
        selectDate(new Date(y, m, d));
      });
    });

    popup.querySelector('.dp-today-btn').addEventListener('click', () => {
      const t = new Date();
      curYear  = t.getFullYear();
      curMonth = t.getMonth();
      selectDate(t);
    });

    popup.querySelector('.dp-clear-btn').addEventListener('click', () => {
      if (target) {
        target.value = '';
        target.dispatchEvent(new Event('input', { bubbles: true }));
        target.dispatchEvent(new Event('change', { bubbles: true }));
      }
      close();
    });

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
    const calBtn = e.target.closest('.cal-btn');
    const calInput = e.target.closest('input.has-cal');

    if (calBtn) {
      e.preventDefault();
      // Найти соответствующий input
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
      // Открываем при клике в поле если ещё не открыт для него
      if (popup && popup.classList.contains('dp-open') && target === calInput) return;
      // Позволим браузеру поставить курсор — не preventDefault
      open(calInput);
      return;
    }

    // Клик вне попапа → закрываем
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
  document.addEventListener('input', function (e) {
    const inp = e.target;
    if (!inp.classList.contains('has-cal')) return;

    let v = inp.value.replace(/[^\d.]/g, '');

    // Добавляем точки автоматически
    if (/^\d{2}$/.test(v)) v += '.';
    else if (/^\d{2}\.\d{2}$/.test(v)) v += '.';

    if (v !== inp.value) inp.value = v;

    // Обновляем подсветку в открытом календаре
    if (popup && popup.classList.contains('dp-open') && target === inp) {
      const parsed = parseDDMMYYYY(v);
      if (parsed) {
        curYear  = parsed.getFullYear();
        curMonth = parsed.getMonth();
      }
      render();
    }
  }, true);

})();
