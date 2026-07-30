# DocGenerator — ГермесГарант

Генератор договоров для риэлторской компании ГермесГарант. Статическое веб-приложение (HTML + CSS + JS), не требует сборки и серверного бэкенда.

## Запуск

```
python3 -m http.server 5000
```

Приложение доступно на порту 5000.

## Стек

- Чистый HTML/CSS/JS (без фреймворков)
- Чтение Excel-файлов сделок через `excel/` (xlsx)
- Генерация Word-документов через `generator/`
- Конфиги агентов и комиссий в `config/`

## Структура

```
index.html          — единственная страница приложения
css/style.css       — все стили
js/                 — логика: app.js, form-builder.js, commission.js и др.
excel/              — чтение и парсинг xlsx-файлов
generator/          — генерация docx-документов
config/             — JSON-конфиги (агенты, комиссии, плейсхолдеры)
assets/             — фото агентов, иконка
fields-config.json  — описание полей формы
```

## User preferences

<!-- Add user preferences here -->
