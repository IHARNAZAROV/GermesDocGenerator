# DocGenerator — Генератор договоров ГермесГарант

## Project overview

An **Electron desktop application** that reads a structured `.xlsx` file and generates Word contracts from templates. The form is **driven entirely by the Excel template** — adding or removing fields in Excel auto-updates the UI without any code changes.

**Stack:** Electron 38, ExcelJS 4, plain HTML/CSS/JS (no frontend framework, no bundler).

## How to run

```
npm install
npm start
```

Requires Node.js with Electron installed (`npm install` handles it via devDependencies).

## How to add / remove fields

1. Edit the Excel template (add or remove rows in column A under a block header).
2. Run the scanner: `node scripts/scan-excel.js <путь/к/шаблону.xlsx>`
3. The scanner updates `fields-config.json` and `js/fields-config.js` automatically.
4. Restart the app — the form reflects the new fields instantly. No other code changes needed.

## Architecture

| File | Role |
|---|---|
| `main.js` | Electron main process — window, IPC handlers |
| `updater.js` | Portable auto-updater — GitHub API check, download, bat/sh replacement script |
| `preload.js` | Context bridge — exposes `window.electronAPI` |
| `index.html` | Renderer HTML — structural skeleton (no hard-coded fields) |
| `js/fields-config.js` | **Auto-generated** — `window.FIELDS_CONFIG` (browser-loadable) |
| `js/form-builder.js` | Builds form HTML dynamically from FIELDS_CONFIG; returns FIELD_MAP |
| `js/app.js` | Renderer logic — uses dynamic FIELD_MAP from form-builder |
| `js/realtor-service.js` | RealtorService — loads `data/realtors.json`, persists selected realtor in localStorage, drives the header dropdown UI |
| `js/money-to-text.js` | BYN → written-out price converter |
| `css/style.css` | All styles — compact CRM theme |
| `excel/excel-reader.js` | Parses `.xlsx` using ExcelJS |
| `generator/word-generator.js` | Fills Word templates via docxtemplater |
| `fields-config.json` | **Single source of truth** — all field definitions, labels, types |
| `scripts/scan-excel.js` | CLI scanner — reads Excel, updates fields-config.json + js/fields-config.js |

## Field config format (`fields-config.json`)

Each field in a group supports:
- `key` — Excel column A value (the field name in the spreadsheet)
- `label` — display label in the UI
- `type` — `"text"` (default), `"date"` (adds calendar button), `"byn"` (price validation), `"computed-propis"` (auto-computed, readonly), `"readonly"`
- `section` — UI section override (e.g. `"deal-prices"` puts deal fields in the property column)
- `pairWith` — renders two inputs in one row (e.g. Корпус/Квартира)
- `pairStyle` — `"slash"` (A / B) or `"floor"` (A из B)
- `pairedUnder` — marks this as the secondary field of a pair
- `computed: true` — field is not read from or written to Excel; preserved through rescans

## Excel structure

One sheet ("Сделка"). Blocks: `СДЕЛКА`, `ОБЪЕКТ`, `ПРОДАВЕЦ`, `СОБСТВЕННИК №1`, `СОБСТВЕННИК №2`, `СОБСТВЕННИК №3`, `ПОКУПАТЕЛЬ`. Column A = field name, Column B = value.

## Replit notes

- **No run workflow** — Electron requires a local desktop environment and cannot run in Replit's browser preview.
- Edit code here; run locally with `npm start`.
- To preview renderer UI changes without Electron, see the potential follow-up task for converting to a web app.

## User preferences

- Do not change `main.js`, `preload.js`, or `excel/excel-reader.js` — these are stable backend/IPC files.
- Only modify the renderer layer: `index.html`, `js/app.js`, `js/form-builder.js`, `css/style.css`, `fields-config.json`.
- Keep the project's existing Electron architecture — do not migrate to a web server unless the user explicitly requests it.
