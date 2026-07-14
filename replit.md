# DocGenerator — Генератор договоров ГермесГарант

## Project overview

An **Electron desktop application** that reads a structured `.xlsx` file and displays its contents as a compact, CRM-style form for reviewing real-estate deal data (сделка, объект, продавец, собственники, покупатель).

**Stack:** Electron 38, ExcelJS 4, plain HTML/CSS/JS (no frontend framework, no bundler).

## How to run

```
npm install
npm start
```

Requires Node.js with Electron installed (`npm install` handles it via devDependencies).

## Architecture

| File | Role |
|---|---|
| `main.js` | Electron main process — creates window, wires IPC handlers |
| `preload.js` | Context bridge — exposes `window.electronAPI` to renderer |
| `index.html` | Renderer HTML — card layout |
| `js/app.js` | Renderer JS — UI logic, form population, collapse, visibility |
| `css/style.css` | All styles — compact CRM theme |
| `excel/excel-reader.js` | Parses `.xlsx` using ExcelJS, returns structured `{deal, property, seller, owner1, owner2, owner3, buyer}` |

## Key UI behaviours

- **Collapsible cards** — each section collapses/expands with a smooth CSS transition. Auto-expanded on load: Сделка, Объект, Покупатель. Others start collapsed.
- **Hide empty fields** — rows with no value are hidden automatically after Excel is loaded.
- **Hide empty sections** — if an entire card (e.g. Собственник №3) has no filled fields, the card is not shown.
- **Sticky toolbar** — the top bar with the "Выбрать Excel" button stays fixed at the top during scroll.

## User preferences

- Do not change `main.js`, `preload.js`, or `excel/excel-reader.js` — these are stable backend/IPC files.
- Only modify the renderer layer: `index.html`, `js/app.js`, `css/style.css`.
- Keep the project's existing Electron architecture — do not migrate to a web server.
