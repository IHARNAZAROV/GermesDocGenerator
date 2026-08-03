'use strict';

// ============================================================
//  name-declension.js — склонение ФИО и организаций
//
//  Поддерживаемые падежи:
//    • Родительный  (Genitive)     — «справка КОГО?»
//    • Дательный    (Dative)       — «принадлежит КОМУ?»
//    • Творительный (Instrumental) — «выдан КЕМ?» (для организаций)
//
//  Все функции глобальные — app.js использует их напрямую.
// ============================================================

// ── Определение пола ─────────────────────────────────────────

function detectGender(middleName, firstName, lastName) {
  const p = (middleName || '').trim().toLowerCase();
  if (p.endsWith('иична') || p.endsWith('овна') || p.endsWith('евна') || p.endsWith('ична')) return 'f';
  if (p.endsWith('ович') || p.endsWith('евич') || p.endsWith('ич')) return 'm';
  const first = (firstName || '').trim().toLowerCase();
  if (first.endsWith('а') || first.endsWith('я') || first.endsWith('ь')) return 'f';
  if (first) return 'm';
  const last = (lastName || '').trim().toLowerCase();
  if (last.endsWith('ова') || last.endsWith('ева') || last.endsWith('ёва') ||
      last.endsWith('ина') || last.endsWith('ына') || last.endsWith('ская') ||
      last.endsWith('цкая') || last.endsWith('зская')) return 'f';
  return null;
}

// ── Творительный падеж для организаций ───────────────────────
//  «выдан Лидским РОВД» вместо «выдан Лидский РОВД»

function toInstrumental(str) {
  if (!str) return str;
  // Длинные/специфичные окончания — первыми, порядок важен.
  return str
    .replace(/(ский)(?=\s|$)/g,  'ским')
    .replace(/(цкий)(?=\s|$)/g,  'цким')
    .replace(/(жний)(?=\s|$)/g,  'жним')
    .replace(/(дний)(?=\s|$)/g,  'дним')
    .replace(/(зний)(?=\s|$)/g,  'зним')
    .replace(/(ний)(?=\s|$)/g,   'ним')
    .replace(/(жный)(?=\s|$)/g,  'жным')
    .replace(/(дный)(?=\s|$)/g,  'дным')
    .replace(/(зный)(?=\s|$)/g,  'зным')
    .replace(/(ный)(?=\s|$)/g,   'ным')
    .replace(/(ий)(?=\s|$)/g,    'им')
    .replace(/(ый)(?=\s|$)/g,    'ым');
}

// ── Дательный падеж (КОМУ?) ───────────────────────────────────

function _declinePatronymicDative(middle) {
  if (!middle) return middle;
  const t = middle.trim();
  const low = t.toLowerCase();
  if (low.endsWith('иична'))  return t.slice(0, -1) + 'е'; // Ильиична → Ильиничне
  if (low.endsWith('овна'))   return t.slice(0, -1) + 'е'; // Казимировна → Казимировне
  if (low.endsWith('евна'))   return t.slice(0, -1) + 'е'; // Андреевна → Андреевне
  if (low.endsWith('ична'))   return t.slice(0, -1) + 'е'; // Ильична → Ильичне
  if (low.endsWith('ович'))   return t + 'у';              // Иванович → Ивановичу
  if (low.endsWith('евич'))   return t + 'у';              // Андреевич → Андреевичу
  if (low.endsWith('ич'))     return t + 'у';              // Ильич → Ильичу
  return t;
}

function _declineFirstNameDative(first, gender) {
  if (!first) return first;
  const t = first.trim();
  const low = t.toLowerCase();
  if (gender === 'f' || gender === null) {
    if (low.endsWith('ия'))  return t.slice(0, -1) + 'и'; // Мария → Марии
    if (low.endsWith('ья'))  return t.slice(0, -1) + 'е'; // Дарья/Наталья → Дарье/Наталье
    if (low.endsWith('я'))   return t.slice(0, -1) + 'е'; // Таня → Тане
    if (low.endsWith('а'))   return t.slice(0, -1) + 'е'; // Юзефа → Юзефе, Анна → Анне
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'и'; // Любовь → Любови
  }
  if (gender === 'm') {
    if (low.endsWith('ий'))  return t.slice(0, -2) + 'ию'; // Василий → Василию
    if (low.endsWith('й'))   return t.slice(0, -1) + 'ю';  // Сергей → Сергею
    if (low.endsWith('я'))   return t.slice(0, -1) + 'е';  // Илья → Илье
    if (low.endsWith('а'))   return t.slice(0, -1) + 'е';  // Никита → Никите
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'ю';  // Игорь → Игорю
    return t + 'у'; // Иван → Ивану, Александр → Александру
  }
  return t;
}

function _declineLastNameDative(last, gender) {
  if (!last) return last;
  const t = last.trim();
  const low = t.toLowerCase();
  if (gender === 'f') {
    if (low.endsWith('ская') || low.endsWith('цкая') || low.endsWith('зская')) {
      return t.slice(0, -2) + 'ой';  // Островская → Островской
    }
    if (low.endsWith('ова') || low.endsWith('ева') || low.endsWith('ёва') ||
        low.endsWith('ина') || low.endsWith('ына')) {
      return t.slice(0, -1) + 'ой';  // Иванова → Ивановой, Пушкина → Пушкиной
    }
    return t; // Фамилия на согласный — не склоняется (Малейкович)
  }
  if (gender === 'm') {
    if (low.endsWith('ский') || low.endsWith('цкий') || low.endsWith('зский')) {
      return t.slice(0, -2) + 'ому'; // Троцкий → Троцкому
    }
    if (low.endsWith('ой') || low.endsWith('ый') || low.endsWith('ий')) {
      return t.slice(0, -2) + 'ому'; // Толстой → Толстому
    }
    if (low.endsWith('ов') || low.endsWith('ев') || low.endsWith('ёв')) {
      return t + 'у'; // Иванов → Иванову
    }
    if (low.endsWith('ин') || low.endsWith('ын')) {
      return t + 'у'; // Пушкин → Пушкину
    }
    if (low.endsWith('ь')) {
      return t.slice(0, -1) + 'ю'; // Медведь → Медведю
    }
    return t + 'у'; // прочие мужские на согласный
  }
  return t;
}

// ── Родительный падеж (КОГО?) ─────────────────────────────────

function _declinePatronymicGenitive(middle) {
  if (!middle) return middle;
  const t = middle.trim();
  const low = t.toLowerCase();
  if (low.endsWith('иична'))  return t.slice(0, -1) + 'ы'; // Ильиична → Ильиичны
  if (low.endsWith('овна'))   return t.slice(0, -1) + 'ы'; // Казимировна → Казимировны
  if (low.endsWith('евна'))   return t.slice(0, -1) + 'ы'; // Андреевна → Андреевны
  if (low.endsWith('ична'))   return t.slice(0, -1) + 'ы'; // Ильична → Ильичны
  if (low.endsWith('ович'))   return t + 'а';              // Иванович → Ивановича
  if (low.endsWith('евич'))   return t + 'а';              // Андреевич → Андреевича
  if (low.endsWith('ич'))     return t + 'а';              // Ильич → Ильича
  return t;
}

function _declineFirstNameGenitive(first, gender) {
  if (!first) return first;
  const t = first.trim();
  const low = t.toLowerCase();
  const velars = ['г', 'к', 'х', 'ж', 'ш', 'щ', 'ч'];
  if (gender === 'f' || gender === null) {
    if (low.endsWith('ия'))  return t.slice(0, -1) + 'и'; // Мария → Марии
    if (low.endsWith('ья'))  return t.slice(0, -2) + 'ьи'; // Дарья → Дарьи
    if (low.endsWith('я'))   return t.slice(0, -1) + 'и'; // Таня → Тани
    if (low.endsWith('а'))   return velars.includes(low[low.length - 2])
      ? t.slice(0, -1) + 'и'   // Ольга → Ольги
      : t.slice(0, -1) + 'ы';  // Юзефа → Юзефы, Анна → Анны
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'и'; // Любовь → Любови
  }
  if (gender === 'm') {
    if (low.endsWith('ий'))  return t.slice(0, -2) + 'ия'; // Василий → Василия
    if (low.endsWith('й'))   return t.slice(0, -1) + 'я';  // Сергей → Сергея
    if (low.endsWith('я'))   return t.slice(0, -1) + 'и';  // Илья → Ильи
    if (low.endsWith('а'))   return velars.includes(low[low.length - 2])
      ? t.slice(0, -1) + 'и'
      : t.slice(0, -1) + 'ы';  // Никита → Никиты
    if (low.endsWith('ь'))   return t.slice(0, -1) + 'я';  // Игорь → Игоря
    return t + 'а'; // Иван → Ивана, Александр → Александра
  }
  return t;
}

function _declineLastNameGenitive(last, gender) {
  if (!last) return last;
  const t = last.trim();
  const low = t.toLowerCase();
  if (gender === 'f') {
    if (low.endsWith('ская') || low.endsWith('цкая') || low.endsWith('зская')) {
      return t.slice(0, -2) + 'ой';  // Островская → Островской
    }
    if (low.endsWith('ова') || low.endsWith('ева') || low.endsWith('ёва') ||
        low.endsWith('ина') || low.endsWith('ына')) {
      return t.slice(0, -1) + 'ой';  // Иванова → Ивановой
    }
    return t; // Фамилия на согласный — не склоняется (Малейкович)
  }
  if (gender === 'm') {
    if (low.endsWith('ский') || low.endsWith('цкий') || low.endsWith('зский')) {
      return t.slice(0, -2) + 'ого'; // Троцкий → Троцкого
    }
    if (low.endsWith('ой') || low.endsWith('ый') || low.endsWith('ий')) {
      return t.slice(0, -2) + 'ого'; // Толстой → Толстого
    }
    if (low.endsWith('ов') || low.endsWith('ев') || low.endsWith('ёв')) {
      return t + 'а'; // Иванов → Иванова
    }
    if (low.endsWith('ин') || low.endsWith('ын')) {
      return t + 'а'; // Пушкин → Пушкина
    }
    if (low.endsWith('ь')) {
      return t.slice(0, -1) + 'я'; // Медведь → Медведя
    }
    return t + 'а'; // прочие мужские на согласный: Горбач → Горбача
  }
  return t;
}

// ── Публичные сборщики ────────────────────────────────────────

function buildNameGenitive(lastName, firstName, middleName) {
  const gender = detectGender(middleName, firstName, lastName);
  const lg = _declineLastNameGenitive(lastName || '', gender);
  const fg = _declineFirstNameGenitive(firstName || '', gender);
  const mg = _declinePatronymicGenitive(middleName || '');
  return {
    lastNameGenitive:   lg,
    firstNameGenitive:  fg,
    middleNameGenitive: mg,
    fullNameGenitive:   [lg, fg, mg].filter(Boolean).join(' '),
  };
}

function buildNameDative(lastName, firstName, middleName) {
  const gender = detectGender(middleName, firstName, lastName);
  const ld = _declineLastNameDative(lastName || '', gender);
  const fd = _declineFirstNameDative(firstName || '', gender);
  const md = _declinePatronymicDative(middleName || '');
  return {
    lastNameDative:   ld,
    firstNameDative:  fd,
    middleNameDative: md,
    fullNameDative:   [ld, fd, md].filter(Boolean).join(' '),
  };
}
