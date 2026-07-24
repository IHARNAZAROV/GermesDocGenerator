# 🔍 Технический аудит проекта DocGenerator (ГермесГарант)

> **Аудит выполнен:** 2026-07-24  
> **Объём проверки:** `js/app.js` (1890 строк), `js/form-builder.js` (484), `js/commission.js` (112), `js/fields-config.js` (516), `js/agents-config.js` (83), `js/datepicker.js` (346), `css/style.css` (3705), `index.html` (858), `excel/excel-reader.js` (88), `excel/excel-scanner.js` (152), `generator/word-generator.js` (105)  
> **Стек:** Electron, Vanilla JS, ExcelJS, Docxtemplater, PizZip, Mammoth  
> **Сценарий:** 8–10 часов в день без перезапуска, несколько риэлторов

---

## 📋 Краткая сводка по приоритетам

| Приоритет | Кол-во | Основные темы |
|-----------|--------|---------------|
| 🔴 Критический | 5 | Накопление EventListener, innerHTML с HTML из IPC, параллельная генерация |
| 🟠 Высокий | 9 | Монолитный app.js, дублирование DOM-запросов, двойная подписка bynInput |
| 🟡 Средний | 11 | DRY-нарушения, дублирование логики склонений, CSS-переменные |
| 🟢 Низкий | 7 | Косметика, console.error в prod, неиспользуемые заглушки |

---

## 🔴 КРИТИЧЕСКИЕ ПРОБЛЕМЫ

---

### 🔴-1 Накопление EventListener при повторном вызове `buildForm` + `_initObjTypeDropdowns`

📁 **Файл:** `js/form-builder.js`, строки 479–480, 388–473  
📖 **Описание:** Внутри `_initObjTypeDropdowns` на каждый вызов `buildForm` добавляются два глобальных слушателя:
```js
document.addEventListener('form:populated', syncAll);
document.addEventListener('form:cleared',   syncAll);
```
Если `buildForm` вызывается несколько раз (например, после «Обновить шаблон»), эти слушатели накапливаются. Через 10 вызовов `syncAll` будет срабатывать 10 раз подряд, опрашивая все `.obj-type-dropdown` каждый раз.

⚠️ **Последствия:** Утечка памяти, задержки при очистке/загрузке формы, труднообнаруживаемые баги.

💡 **Исправление:**
```js
// Добавить флаг инициализации
let _globalListenersAttached = false;
function _initObjTypeDropdowns() {
  // ... инициализация дропдаунов ...
  if (!_globalListenersAttached) {
    document.addEventListener('form:populated', syncAll);
    document.addEventListener('form:cleared',   syncAll);
    _globalListenersAttached = true;
  }
}
```

📈 **Эффект:** Устранение утечки, стабильная работа после «Обновить шаблон».

---

### 🔴-2 XSS через `innerHTML` с данными из IPC (превью документа)

📁 **Файл:** `js/app.js`, строки 1714–1718  
📖 **Описание:** HTML, полученный от `electronAPI.previewDocument()` через IPC, напрямую вставляется в DOM через `innerHTML`:
```js
const page = document.createElement('div');
page.className = 'preview-page';
page.innerHTML = result.html;  // ← данные из main-процесса, сгенерированные mammoth
```
`mammoth` конвертирует `.docx` → HTML. Если шаблон Word содержит вредоносный HTML (тег `<script>`, `<img onerror=...>`), он попадёт в renderer. CSP заголовок (`script-src 'self'`) блокирует `<script>`, но `onerror` на изображениях — нет.

⚠️ **Последствия:** Потенциальный XSS при использовании сторонних/подменённых шаблонов. В Electron renderer с `nodeIntegration: false` риск ограничен, но при изменении настроек BrowserWindow может дать полный доступ к Node.

💡 **Исправление:**
```js
// Вариант 1: санитизация через DOMParser (без внешних зависимостей)
const parser = new DOMParser();
const doc = parser.parseFromString(result.html, 'text/html');
// Удалить опасные атрибуты
doc.querySelectorAll('[onerror],[onload],[onclick]').forEach(el => {
  ['onerror','onload','onclick','onmouseover'].forEach(attr => el.removeAttribute(attr));
});
page.appendChild(doc.body);

// Вариант 2: использовать DOMPurify (добавить как зависимость)
page.innerHTML = DOMPurify.sanitize(result.html);
```

📈 **Эффект:** Защита от XSS через вредоносные Word-шаблоны.

---

### 🔴-3 Параллельная генерация документов через `await` в цикле — блокировка UI

📁 **Файл:** `js/app.js`, строки 1651–1674  
📖 **Описание:** Генерация нескольких документов выполняется последовательно через `await` в цикле `for...of`:
```js
for (const key of toGenerate) {
  const result = await entry.generate(outputDir, options); // последовательно!
}
```
При выборе 5+ шаблонов каждый генерируется по очереди: главный поток ждёт завершения каждого IPC-вызова. Генерация одного `.docx` через Docxtemplater может занимать 200–800 мс.

⚠️ **Последствия:** При 10 шаблонах — задержка 2–8 секунд без индикатора прогресса. UI «висит».

💡 **Исправление:**
```js
// Параллельная генерация с Promise.allSettled
const results = await Promise.allSettled(
  toGenerate.map(key => TEMPLATE_REGISTRY[key].generate(outputDir, options))
);
results.forEach((r, i) => {
  if (r.status === 'fulfilled' && r.value?.success) {
    successCount++;
    // ... история
  } else {
    errors.push(`${TEMPLATE_REGISTRY[toGenerate[i]].label}: ${r.reason?.message || r.value?.error}`);
  }
});
```

📈 **Эффект:** Ускорение генерации в 3–5× при нескольких шаблонах.

---

### 🔴-4 `autoUpdateCommission` вызывает `calculateCommission` синхронно до загрузки JSON-тарифов

📁 **Файл:** `js/commission.js`, строки 15–19, 84–86; `js/app.js`, строки 636–651  
📖 **Описание:** `window.COMMISSION_CONFIG` инициализируется с `brackets: []`, промис `COMMISSION_CONFIG_READY` загружает данные асинхронно. Если `autoUpdateCommission()` вызывается до разрешения промиса (при быстром вводе цены сразу после старта), `calculateCommission` получает пустой массив `brackets`:
```js
const bracket = bkts.find(b => b.upTo === null || baseUnits <= b.upTo);
// bkts = [] → bracket = undefined
const percent = bracket ? bracket.percent : 1.5; // всегда 1.5%!
```

⚠️ **Последствия:** Некорректный расчёт комиссии (всегда 1.5%) при быстром вводе сразу после запуска. Молчащий баг — пользователь не узнает об ошибке.

💡 **Исправление:**
```js
// В app.js — ждать готовности тарифов перед подпиской
window.COMMISSION_CONFIG_READY.then(() => {
  if (bynInput) bynInput.addEventListener('input', autoUpdateCommission);
  autoUpdateCommission(); // пересчитать если уже есть значение
});
```

📈 **Эффект:** Гарантированно корректный расчёт комиссии с первого запуска.

---

### 🔴-5 `render()` в datepicker вызывает `popup.innerHTML = html` при каждом нажатии клавиши

📁 **Файл:** `js/datepicker.js`, строки 323–344, 127–202  
📖 **Описание:** При ручном вводе даты (поле `has-cal`) каждый символ вызывает `render()` через глобальный `input`-слушатель:
```js
document.addEventListener('input', function (e) {
  // ...
  render(); // ← полная перестройка DOM попапа при каждом символе!
});
```
`render()` полностью перезаписывает `popup.innerHTML` (~60+ элементов календаря), добавляет ~40 новых EventListener, выполняет layout-пересчёт для позиционирования. Это происходит на каждое нажатие клавиши.

⚠️ **Последствия:** При быстром вводе даты — микрозависания, лишний reflow, накопление задержки за день работы.

💡 **Исправление:**
```js
// Дебаунс на перерисовку + только обновление выделения без полного перерендера
let _renderTimer = null;
document.addEventListener('input', function (e) {
  const inp = e.target;
  if (!inp.classList.contains('has-cal')) return;
  // ... автодоточки ...
  if (popup?.classList.contains('dp-open') && target === inp) {
    clearTimeout(_renderTimer);
    _renderTimer = setTimeout(() => {
      const parsed = parseDDMMYYYY(inp.value);
      if (parsed) { curYear = parsed.getFullYear(); curMonth = parsed.getMonth(); }
      render();
    }, 150);
  }
}, true);
```

📈 **Эффект:** Снижение нагрузки на DOM при вводе дат, плавный UX.

---

## 🟠 ВЫСОКИЙ ПРИОРИТЕТ

---

### 🟠-1 Монолитный `app.js` — 1890 строк, смешение 7 зон ответственности

📁 **Файл:** `js/app.js`  
📖 **Описание:** Файл содержит: управление файлами Excel, грязное состояние, автосохранение, валидацию, склонение имён (дательный/родительный падежи), сборку данных шаблонов, регистр шаблонов, генерацию документов, превью, модальные окна, автообновление — всё в одном файле.

⚠️ **Последствия:** Сложность поддержки, риск случайного нарушения несвязанной функциональности, трудность написания тестов.

💡 **Рекомендуемое разбиение:**
```
js/
  state-manager.js     — dirtyInputIds, originalValues, autoSave
  excel-handler.js     — loadExcelFile, handleSave, handleSaveAs, buildUpdates
  form-populator.js    — populateForm, clearAllInputs, setInputValue
  name-declension.js   — buildNameGenitive, buildNameDative, toInstrumental (~300 строк)
  template-data.js     — buildPlaceholderData, buildPersonBlock, GENITIVE_MAP
  document-gen.js      — TEMPLATE_REGISTRY, handleGenerate, openPreviewModal
  app.js               — только инициализация и связывание модулей (~150 строк)
```

📈 **Эффект:** Снижение когнитивной нагрузки, изолированное тестирование, сокращение app.js до ~200 строк.

---

### 🟠-2 Двойная подписка `bynInput` на `input`

📁 **Файл:** `js/app.js`, строки 629–631 и 653–655  
📖 **Описание:**
```js
if (bynInput) {
  bynInput.addEventListener('input', autoUpdatePropis);     // строка 630
}
// ...
if (bynInput) {
  bynInput.addEventListener('input', autoUpdateCommission); // строка 654
}
```
Обработчиков два — это нормально, но оба работают независимо и выполняют схожий парсинг `bynInput.value`. Кроме того, `autoUpdateCommission` не ожидает `COMMISSION_CONFIG_READY` (см. 🔴-4).

💡 **Исправление:**
```js
if (bynInput) {
  bynInput.addEventListener('input', () => {
    autoUpdatePropis();
    autoUpdateCommission();
  });
}
```
📈 **Эффект:** Один проход по значению поля вместо двух, ясный порядок обновления.

---

### 🟠-3 `onInputChange` выполняет `document.getElementById(inputId)` повторно

📁 **Файл:** `js/app.js`, строки 187–199  
📖 **Описание:** Функция `onInputChange` принимает `inputId`, но внутри снова делает `document.getElementById(inputId)` — хотя `e.target` уже доступен на месте вызова:
```js
// Вызов (строка 243):
onInputChange(id, e.target.value);

// Внутри функции (строка 191):
const el = document.getElementById(inputId); // повторный поиск!
```

💡 **Исправление:** Передавать элемент напрямую:
```js
function onInputChange(el, currentValue) { ... }
// Вызов:
onInputChange(e.target, e.target.value);
```

📈 **Эффект:** Устранение лишнего DOM-запроса при каждом вводе (срабатывает при каждом нажатии клавиши).

---

### 🟠-4 `applyObjectTypeVisibility` и `updateContractAvailability` — дорогие querySelectorAll при каждом вводе

📁 **Файл:** `js/app.js`, строки 871–892, 837–854  
📖 **Описание:** Обе функции вызываются из `onInputChange` при изменении определённых полей и каждый раз обходят весь DOM:
```js
document.querySelectorAll('[data-object-type]')          // все поля с типом объекта
document.querySelectorAll('.tpl-item[data-owners-required]') // все пункты шаблонов
```
При 500 шаблонах `updateContractAvailability` будет обходить 500 элементов на каждое изменение поля собственника.

💡 **Исправление:** Кэшировать NodeList при инициализации:
```js
// Один раз при старте
const _objectTypeEls = [...document.querySelectorAll('[data-object-type]')];
const _tplOwnerItems = [...document.querySelectorAll('.tpl-item[data-owners-required]')];

// В функциях — итерировать по кэшированным массивам
```

📈 **Эффект:** Снижение нагрузки на DOM при работе с полями, особенно критично при большом количестве шаблонов.

---

### 🟠-5 `updateBlockCompletion` → `getRequiredIssueCountByBlock()` → `getIssues()` при каждом вводе

📁 **Файл:** `js/app.js`, строки 779–826, 729–735  
📖 **Описание:** При каждом `input` в любом поле формы вызывается `updateBlockCompletion`, которая вызывает `getRequiredIssueCountByBlock()`, которая вызывает глобальную `getIssues()` из `ui-controller.js`. Если `getIssues()` делает DOM-обход (скорее всего), это происходит при каждом нажатии клавиши.

💡 **Исправление:** Дебаунс для `updateBlockCompletion`:
```js
let _blockCompletionTimer = null;
function scheduleBlockCompletion(prefix) {
  clearTimeout(_blockCompletionTimer);
  _blockCompletionTimer = setTimeout(() => updateBlockCompletion(prefix), 200);
}
// В onInputChange — вызывать scheduleBlockCompletion вместо updateBlockCompletion
```

📈 **Эффект:** Сокращение вызовов `getIssues()` в 10–20× при быстром вводе.

---

### 🟠-6 `handleGenerate` — `await` внутри цикла = последовательное выполнение (повтор с другого угла)

📁 **Файл:** `js/app.js`, строки 1651–1674  
📖 **Описание:** (Связано с 🔴-3) Помимо производительности, при ошибке в одном документе цикл **продолжается**, но без `showLoader()` — пользователь не видит прогресс генерации множества файлов.

💡 **Дополнительное исправление:** Показывать `showLoader()` на время генерации и скрывать после завершения:
```js
showLoader();
try {
  // Promise.allSettled(...)
} finally {
  hideLoader();
}
```

---

### 🟠-7 `cellToString` и `cellText` — идентичная логика в двух файлах

📁 **Файлы:** `excel/excel-reader.js` (строки 13–32), `excel/excel-scanner.js` (строки 17–29)  
📖 **Описание:** Обе функции делают одно и то же — конвертируют ячейку ExcelJS в строку, обрабатывая `Date`, `richText`, `result`, `text`. Код практически идентичен, но функции называются по-разному и живут в разных файлах.

💡 **Исправление:** Вынести в общий модуль:
```js
// excel/cell-utils.js
function cellToString(cell) { ... }
module.exports = { cellToString };

// В обоих файлах:
const { cellToString } = require('./cell-utils');
```

📈 **Эффект:** ~15 строк дублирования убрано, единая точка исправления при изменении формата Excel.

---

### 🟠-8 `inferType` в `excel-scanner.js` — только проверка на "дата", все остальное "text"

📁 **Файл:** `excel/excel-scanner.js`, строки 31–33  
📖 **Описание:**
```js
function inferType(key) {
  return key.toLowerCase().startsWith('дата') ? 'date' : 'text';
}
```
Новые поля из Excel всегда получают тип `text`, даже если ключ содержит «цена», «стоимость», «сумма» (должен быть `byn` или `numeric`).

⚠️ **Последствия:** После сканирования нового Excel числовые поля не получают форматирование тысяч, BYN-поля не вычисляют сумму прописью — пользователь не замечает, данные записываются некорректно.

💡 **Исправление:**
```js
function inferType(key) {
  const k = key.toLowerCase();
  if (k.startsWith('дата')) return 'date';
  if (k.includes('byn') || k.includes('бyn') || k.includes('сумма') || k.includes('стоимость')) return 'byn';
  if (k.includes('usd') || k.includes('площадь') || k.includes('количество')) {
    // добавить numeric: true в объект поля
  }
  return 'text';
}
```

---

### 🟠-9 `buildPlaceholderData` — `getField('...')` вызывается 40+ раз, каждый раз делает `document.getElementById`

📁 **Файл:** `js/app.js`, строки 1408–1581  
📖 **Описание:** `buildPlaceholderData()` вызывает `getField()` более 40 раз. Каждый `getField(id)` выполняет `document.getElementById(id)`. Хотя это быстрая операция, при вызове через кнопку «Генерировать» или «Превью» получаем 40+ DOM-обращений синхронно. Некоторые поля запрашиваются дважды (например, `'deal-Номер договора'` → `deal.number` и `deal.contractNumber` — одно и то же значение).

💡 **Исправление:**
```js
// Снэпшот всех значений за один проход
function snapshotFields() {
  const snap = {};
  for (const id of FIELD_IDS) {
    const el = document.getElementById(id);
    if (el) snap[id] = el.dataset.numeric ? el.value.replace(/\s/g,'').trim() : el.value.trim();
  }
  return snap;
}
```
Передавать снэпшот в `buildPlaceholderData(snap)`.

📈 **Эффект:** 1 проход по DOM вместо 40+, более предсказуемое поведение.

---

## 🟡 СРЕДНИЙ ПРИОРИТЕТ

---

### 🟡-1 Дублирование структуры `owner1`, `owner2`, `owner3` — 36 повторяющихся полей

📁 **Файл:** `js/fields-config.js`, строки 292–461  
📖 **Описание:** Три группы `owner1`, `owner2`, `owner3` содержат идентичные 12 полей каждая. Итого ~36 строк JSON с одинаковыми ключами `Фамилия`, `Имя`, `Отчество` и т.д. Если нужно добавить поле или изменить метку — придётся менять в трёх местах.

💡 **Исправление:** Генерировать при инициализации:
```js
const OWNER_FIELDS_TEMPLATE = [
  { key: 'Фамилия', label: 'Фамилия:' },
  // ...
];
const ownerGroups = [1,2,3].map(n => ({
  id: `owner${n}`,
  excelHeader: `СОБСТВЕННИК №${n}`,
  defaultSection: `owner${n}`,
  fields: [...OWNER_FIELDS_TEMPLATE, { key: 'Доля собственности', label: 'Доля собств.:' }],
}));
```

📈 **Эффект:** ~80 строк убрано, единая точка изменения структуры собственников.

---

### 🟡-2 Определение пола в `buildNameGenitive` и `buildNameDative` — полное дублирование

📁 **Файл:** `js/app.js`, строки 1336–1373  
📖 **Описание:** Обе функции содержат идентичный блок определения пола по отчеству:
```js
// В buildNameGenitive (строки 1337–1342) — ИДЕНТИЧНО buildNameDative (1356–1361)
const p = (middleName || '').trim().toLowerCase();
let gender = null;
if (p.endsWith('иична') || p.endsWith('овна') || ...) gender = 'f';
else if (p.endsWith('ович') || ...) gender = 'm';
```

💡 **Исправление:**
```js
function detectGender(middleName) {
  const p = (middleName || '').trim().toLowerCase();
  if (p.endsWith('иична') || p.endsWith('овна') || p.endsWith('евна') || p.endsWith('ична')) return 'f';
  if (p.endsWith('ович') || p.endsWith('евич') || p.endsWith('ич')) return 'm';
  return null;
}
function buildNameGenitive(last, first, middle) {
  const gender = detectGender(middle);
  // ...
}
```
📈 **Эффект:** ~8 строк убрано, единая логика определения пола.

---

### 🟡-3 `buildPersonBlock` вызывает `getField` + склонения для каждого `buildPlaceholderData`

📁 **Файл:** `js/app.js`, строки 1374–1406  
📖 **Описание:** `buildPersonBlock` вычисляет все падежи имён при каждом вызове `buildPlaceholderData`. Каждый вызов `buildNameGenitive` + `buildNameDative` — это 6+ вызовов вспомогательных функций с regex-операциями. Вызывается для 6 блоков (seller, owner1-3, buyer) при каждом открытии превью.

💡 **Рекомендация:** Мемоизировать результаты для неизменённых блоков через `originalValues` как ключ кэша.

---

### 🟡-4 Дублирование JSDoc для `calculateCommission`

📁 **Файл:** `js/commission.js`, строки 71–83  
📖 **Описание:** Функция `calculateCommission` имеет два JSDoc-блока подряд — старый (строки 71–76) и новый расширенный (строки 77–83). Первый устарел и вводит в заблуждение.

💡 **Исправление:** Удалить строки 71–76, оставить только актуальный JSDoc.

---

### 🟡-5 `FALLBACK_BRACKETS` в `commission.js` — повторяется для коммерческой недвижимости

📁 **Файл:** `js/commission.js`, строки 23–39, 67  
📖 **Описание:** При ошибке загрузки `commission-rates-commercial.json` используется тот же `FALLBACK_BRACKETS` что и для стандартной таблицы. Это скрытый баг: коммерческая комиссия может быть принципиально другой, а пользователь не узнает, что применились стандартные тарифы.

💡 **Исправление:** Отдельный `FALLBACK_COMMERCIAL_BRACKETS` или явное предупреждение:
```js
console.warn('[commission] Используется запасная таблица для коммерческой недвижимости!');
```

---

### 🟡-6 `style.css` — 3705 строк в одном файле, hardcoded цвета вне CSS-переменных

📁 **Файл:** `css/style.css`  
📖 **Описание:** Файл содержит все стили приложения. Ряд значений продублированы без переменных:
- `#EF4444` (red) — встречается 4+ раз вместо переменной
- `rgba(21,89,69,...)` — встречается 10+ раз в разных непоследовательных значениях opacity
- `#B45309` (dirty) — встречается 1 раз в vars, но `--dirty-bdr: #F5A623` жёстко прописан
- Переходы `.15s` повторяются ~30 раз

💡 **Исправление:**
1. Добавить CSS-переменные: `--red: #EF4444`, `--brand-rgb: 21,89,69`
2. Разбить на файлы: `base.css`, `toolbar.css`, `form.css`, `datepicker.css`, `modals.css`

---

### 🟡-7 `backdrop-filter: blur(6px)` на loader-overlay — GPU-нагрузка

📁 **Файл:** `css/style.css`, строка 169  
📖 **Описание:** Loader overlay использует `backdrop-filter: blur(6px)`. Каждый раз при показе/скрытии лоадера (при каждом чтении файла) GPU перерисовывает всю область экрана с размытием.

💡 **Исправление:** Заменить на полупрозрачный фон без blur:
```css
.loader-overlay {
  background: rgba(246,248,247,.92); /* без backdrop-filter */
}
```
Или использовать `will-change: transform` на лоадере для GPU-слоя.

---

### 🟡-8 `void section.offsetWidth` — принудительный reflow для анимации

📁 **Файл:** `js/app.js`, строка 818  
📖 **Описание:** Для перезапуска CSS-анимации используется принудительный reflow:
```js
section.classList.remove('ws-block--complete-flash');
void section.offsetWidth; // ← принудительный layout
section.classList.add('ws-block--complete-flash');
```
Срабатывает при каждом завершении блока.

💡 **Альтернатива:**
```js
// Через animation name toggle — без reflow
section.style.animationName = 'none';
requestAnimationFrame(() => {
  section.style.animationName = '';
  section.classList.add('ws-block--complete-flash');
});
```

---

### 🟡-9 `renderFields` + `renderFieldsTwoCol` — оба создают `byKey` lookup

📁 **Файл:** `js/form-builder.js`, строки 151–152, 217–219  
📖 **Описание:** Обе функции создают `byKey = {}` и итерируют `fields.forEach(f => { byKey[f.key] = f; })`. При вызове `renderFieldsTwoCol` → `renderFields` lookup создаётся дважды для одного набора полей.

💡 **Исправление:** Передавать `byKey` параметром или создавать один раз перед вызовом.

---

### 🟡-10 `handleCheckData` — хардкод полей покупателя в 10-элементном массиве

📁 **Файл:** `js/app.js`, строки 545–557  
📖 **Описание:** Проверка полей покупателя захардкодена массивом из 10 пар `[id, label]`. Если структура покупателя изменится в `fields-config.js` — `handleCheckData` нужно обновлять вручную.

💡 **Исправление:** Генерировать список полей из `FIELDS_CONFIG.groups` динамически, как это сделано в `ownerCheckFields`.

---

### 🟡-11 `excel-reader.js` — читает только первый лист (`workbook.worksheets[0]`)

📁 **Файл:** `excel/excel-reader.js`, строка 38  
📖 **Описание:** Если файл Excel содержит несколько листов или данные не на первом листе — чтение вернёт пустой результат без предупреждения.

💡 **Исправление:** Добавить поиск листа по имени или предупреждать если нашли несколько:
```js
const worksheet = workbook.worksheets.find(ws =>
  ws.name.includes('Сделка') || ws.name.includes('Deal')
) || workbook.worksheets[0];
```

---

## 🟢 НИЗКИЙ ПРИОРИТЕТ

---

### 🟢-1 Trailing comma в `agents-config.js`

📁 **Файл:** `js/agents-config.js`, строка 42  
📖 **Описание:** После записи `turko` стоит trailing comma + пустая строка (строки 42–44), что нарушает единообразие форматирования.

---

### 🟢-2 `console.error` в `commission.js` — попадает в production

📁 **Файл:** `js/commission.js`, строки 54, 65  
📖 **Описание:** `console.error(...)` при ошибке загрузки JSON оставляет сообщения в консоли Electron для конечных пользователей.

💡 **Рекомендация:** Использовать `logger` с уровнями или показывать уведомление через `showToast` в renderer.

---

### 🟢-3 `agency` — полностью пустой объект в `buildPlaceholderData`

📁 **Файл:** `js/app.js`, строки 1534–1538  
📖 **Описание:** Объект `agency` создаётся с пустыми полями и никогда не заполняется:
```js
const agency = {
  name: '', shortName: '', director: '', address: '',
  phone: '', email: '', website: '', bank: '',
  bankAccount: '', bik: '', unp: '',
};
```
Если шаблон `.docx` содержит `{{agency.name}}` — он будет пустым.

💡 **Рекомендация:** Добавить `agency.json` в `config/` и загружать аналогично `agents.json`.

---

### 🟢-4 `keys` и `money` — заглушки-пустышки

📁 **Файл:** `js/app.js`, строки 1540–1541  
📖 **Описание:** Аналогично `agency` — `keys` и `money` никогда не заполняются реальными данными.

---

### 🟢-5 `deal.dateText` — всегда пустая строка

📁 **Файл:** `js/app.js`, строка 1418  
📖 **Описание:** `deal.dateText: ''` — поле объявлено, но никогда не заполняется. При этом `deal.endDateText` заполняется через `dateToLongRussian`. Вероятно, забыли добавить: `deal.dateText: dateToLongRussian(getField('deal-Дата договора'))`.

---

### 🟢-6 Дублирование: `deal.number` и `deal.contractNumber` — одно и то же значение

📁 **Файл:** `js/app.js`, строки 1416, 1421  
📖 **Описание:**
```js
number:         getField('deal-Номер договора') || '',
// ...
contractNumber: getField('deal-Номер договора') || '', // дубль
```

---

### 🟢-7 `property.priceUSD` находится внутри объекта `property`, а не `deal`

📁 **Файл:** `js/app.js`, строки 1460–1462  
📖 **Описание:** `priceUSD`, `priceBYN`, `priceWords` логически относятся к сделке (`deal`), но помещены в объект `property`. Это нарушает семантику для разработчиков, создающих новые шаблоны.

---

## 📊 Итоговая таблица

| # | Файл | Строки | Приоритет | Тема |
|---|------|--------|-----------|------|
| 🔴-1 | form-builder.js | 479–480 | Критический | Накопление EventListener |
| 🔴-2 | app.js | 1714–1718 | Критический | XSS через innerHTML из IPC |
| 🔴-3 | app.js | 1651–1674 | Критический | await в цикле, блокировка UI |
| 🔴-4 | commission.js | 15–86 | Критический | Расчёт до загрузки JSON |
| 🔴-5 | datepicker.js | 323–344 | Критический | render() на каждый keystroke |
| 🟠-1 | app.js | все 1890 | Высокий | Монолит, разбить на модули |
| 🟠-2 | app.js | 629, 653 | Высокий | Двойная подписка bynInput |
| 🟠-3 | app.js | 187–199 | Высокий | Повторный getElementById |
| 🟠-4 | app.js | 871, 837 | Высокий | querySelectorAll при каждом input |
| 🟠-5 | app.js | 779–826 | Высокий | getIssues() при каждом input |
| 🟠-6 | app.js | 1651 | Высокий | Нет loader при генерации |
| 🟠-7 | excel/*.js | 13, 17 | Высокий | Дублирование cellToString |
| 🟠-8 | excel-scanner.js | 31 | Высокий | inferType только для дат |
| 🟠-9 | app.js | 1408–1581 | Высокий | 40+ getElementById в buildPlaceholderData |
| 🟡-1 | fields-config.js | 292–461 | Средний | 36 дублированных полей собственников |
| 🟡-2 | app.js | 1337, 1357 | Средний | Дубль определения пола |
| 🟡-3 | app.js | 1374 | Средний | Склонения при каждом превью |
| 🟡-4 | commission.js | 71–76 | Средний | Дублированный JSDoc |
| 🟡-5 | commission.js | 67 | Средний | Неверный коммерческий fallback |
| 🟡-6 | style.css | весь файл | Средний | 3705 строк, hardcoded цвета |
| 🟡-7 | style.css | 169 | Средний | backdrop-filter: blur |
| 🟡-8 | app.js | 818 | Средний | void offsetWidth reflow |
| 🟡-9 | form-builder.js | 151, 217 | Средний | Двойной byKey lookup |
| 🟡-10 | app.js | 545 | Средний | Хардкод полей покупателя |
| 🟡-11 | excel-reader.js | 38 | Средний | Только первый лист |
| 🟢-1 | agents-config.js | 42 | Низкий | Trailing comma |
| 🟢-2 | commission.js | 54, 65 | Низкий | console.error в production |
| 🟢-3 | app.js | 1534 | Низкий | Пустой agency-объект |
| 🟢-4 | app.js | 1540 | Низкий | Пустые keys/money |
| 🟢-5 | app.js | 1418 | Низкий | deal.dateText всегда '' |
| 🟢-6 | app.js | 1416, 1421 | Низкий | deal.number === deal.contractNumber |
| 🟢-7 | app.js | 1460 | Низкий | priceUSD в объекте property |

---

## 🚀 Оценка масштабируемости

| Сценарий | Узкое место | Оценка |
|----------|-------------|--------|
| 500 шаблонов | `updateContractAvailability` — querySelectorAll 500 элементов на каждый input | ⚠️ Заметно |
| 20 000 объектов | Не влияет — загружается по одному | ✅ Нет проблем |
| 10 000 сделок | `RecentDocs` в localStorage — ограничен 5MB | ⚠️ Ограничение |
| 100 открытых документов | Mammoth конвертирует в HTML синхронно в Node — потенциальная очередь IPC | ⚠️ Заметно |
| Работа весь день без перезапуска | Накопление EventListener (🔴-1) + reflow datepicker (🔴-5) | 🔴 Критично |

---

## 🏆 Рекомендуемый порядок исправлений

1. **Немедленно:** 🔴-1 (накопление ListenerS), 🔴-4 (комиссия до загрузки), 🔴-5 (datepicker дебаунс)  
2. **На ближайшем спринте:** 🔴-2 (XSS preview), 🔴-3 (параллельная генерация), 🟠-2, 🟠-3, 🟠-4, 🟠-5  
3. **При рефакторинге:** 🟠-1 (разбивка app.js), 🟠-7 (общий cellToString), 🟡-1 (дублирование owners), 🟡-6 (CSS-модули)  
4. **Когда удобно:** Все 🟢 пункты
