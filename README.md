# ReverseBill (مراجع الفواتير)

Desktop web app for reviewing supplier invoices against approved price lists pulled from a single published Google Sheets document (one tab per list).

## Features

- Load price lists from a published Google Sheets file via CSV (5 lists: AgentDist, Company, online, Retail, Shaheen) with automatic gid discovery
- In-memory list store (`window.appLists`) with fast indexed lookup
- Invoice paste from clipboard or Excel import, automatic price matching against the selected list
- Variance analysis per line item (unit diff, total diff, status badges: match / high / save / unknown)
- Print report with repeatable table header, summary cards, and signature blocks
- Local-only storage (no server required, all data stays in browser localStorage)

## Tech Stack

- Vanilla HTML / CSS / JavaScript (no build step, no framework)
- Tailwind CSS via CDN
- Google Fonts (Tajawal)
- Font Awesome icons via CDN
- Sheets.js wrapper for Google Sheets CSV API

## Setup

1. Open `index.html` in any modern browser — no server required.
2. Paste a published Google Sheets URL (or use the default included in `config.js`).
3. Match list tabs to app list slugs via the Data / Cloud tab.

## Project Structure

```
css/style.css          – App styles + print media rules
js/config.js           – Configuration constants
js/utils.js            – Normalization and formatting helpers
js/storage.js          – localStorage persistence
js/sheets.js           – Google Sheets published CSV fetcher
js/parser.js           – Clipboard / TSV / CSV row parser
js/lists.js            – Price list CRUD, in-memory store, indexed matching
js/comparison.js       – Variance analysis and aggregation
js/invoices.js         – Invoice CRUD, draft persistence, upsert save
js/state.js            – Central state and pub/sub bus
js/app.js              – UI rendering and event binding
js/excel.js            – Excel export (SheetJS)
index.html             – Single-page app shell
```
