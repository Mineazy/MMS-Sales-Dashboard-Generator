# MMS Sales Dashboard - Project Summary

## Overview
Built an interactive sales dashboard that pulls data from a Kobo Toolbox project (MMS Sales Tracker) with 18,586+ records spanning March–July 2026. The dashboard is served by an Express.js server and includes metrics, charts, filters, search, data export capabilities, and multi-rep invoice handling.

## Architecture

### Files
- **`server.js`** — Express server (port 3001) that:
  - Serves the dashboard HTML at `GET /`
  - Provides filtered data API at `GET /api/data`
  - Handles data refresh from Kobo at `GET /api/refresh`
  - Generates Excel exports at `GET /api/export/xlsx`
  - All filtering and aggregation happens server-side
- **`kobo_data.xlsx`** — Cached data file downloaded from Kobo Toolbox
- **`package.json`** — Dependencies: `express`, `cors`, `xlsx`
- **`MEMORY.md`** — This project summary
- **`dashboard.html`** — Initial static version (superseded by server)
- **`generate_dashboard.js`** — Initial generation script (superseded by server)

### Data Source
Kobo Toolbox API endpoint: `/api/v2/assets/ai5wFipp7kxWNA5RJd5jre/export-settings/esdqM5MrXnavHqJcgWRZwcc/data.xlsx`

## Features Implemented

### 1. Core Dashboard (Round 1)
- **Metric Cards**: Total Sales ($406M+), Total Invoices (18,586+), Active Sales Reps (173), Avg Daily Sales, Best Performer, Avg Invoice
- **Charts**: Monthly Sales Trend (bar), Top 20 Sales Reps (horizontal bar), Daily Sales (line)
- **Rep Performance Table**: Ranked list with sales totals, invoice counts, avg invoice, % of invoices, visual bars
- **Transaction Table**: Up to 500 records with date, invoice, reps, branch, amount

### 2. Search & Filters (Round 2)
- **Search**: Real-time search across invoice numbers, rep labels, branches, dates
- **Period filter**: Date range picker (From / To)
- **Month filter**: Dropdown of available months
- **Sales Rep filter**: Dropdown of all active reps
- **Invoice filter**: Text search for specific invoice numbers
- **Reset button**: Clears all filters at once
- **All filters update charts, metrics, and tables dynamically** via API calls

### 3. Data Refresh (Round 2 + Round 6 Fix)
- **"Pull Latest Data" button**: Downloads fresh data from the Kobo API endpoint
- Data is cached to `kobo_data.xlsx` for persistence across server restarts
- On startup, loads from cache if available; falls back to fetching from Kobo
- **Round 6 fix**: Kobo API returns 302 redirects to relative paths (`/private-media/...`). Added `resolveUrl()` helper to combine relative redirects with the original domain, plus redirect depth limit (max 5) and HTTP status error handling

### 4. Top 20 Reps by Month (Round 3)
- Separate month dropdown for the Rep Performance section (chart + table)
- Both dropdowns are synced (chart and table share the same selection)
- Updates the "Best Performer" metric card dynamically
- Defaults to "All Time" showing overall rankings

### 5. PDF Export (Round 3)
- Uses jsPDF + autoTable (loaded from CDN)
- Generates a professional PDF report from the currently filtered rep data
- Includes title, date range summary, and formatted rep table
- Respects all active filters

### 6. Excel Export (Round 3)
- Server-side generation using the `xlsx` library
- Multi-sheet workbook: Summary, Top Reps, Transactions, Monthly, Daily
- Triggered via `GET /api/export/xlsx?filter=params`
- Respects all active filters

### 7. Multi-Rep Invoice Handling (Round 4 & 5)
- **Field**: `Sales Rep Number(s)` contains additional rep numbers when multiple reps worked on an invoice
- **Separators**: `/`, `;`, `\` (all three are handled)
- **Logic**: Each rep gets full credit for the invoice amount
- `extractReps()` function parses each row: checks `Sales Rep Number(s)` first, falls back to `Sales Rep Number`
- Records store `reps` array and `repLabel` for display (e.g., "Rep #32, #09")
- Rep filtering matches any rep in the multi-rep list
- Impact: Rep #20 went from 2,974 → 3,101 records (+127 from multi-rep invoices)

### 8. Top 20 Ranked by Invoice Count (Round 6)
- Changed ranking from sales amount to invoice count
- `repSorted` sorts by `count` (invoices) instead of `amount` (sales)
- Percentage column shows `% of Invoices` instead of `% of Sales`
- Both `processData` and `computeFiltered` use the same sort order
- Best Performer card correctly shows the top rep by invoice count

## API Endpoints

| Endpoint | Description |
|---|---|
| `GET /` | Serves the dashboard HTML |
| `GET /api/data?search=&month=&rep=&invoice=&dateFrom=&dateTo=&top=` | Returns filtered data as JSON |
| `GET /api/export/xlsx?...` | Downloads filtered data as Excel workbook |
| `GET /api/refresh` | Pulls fresh data from Kobo and caches it |

## How to Run

```bash
cd "C:\Users\Administrator\Sales Dashboard Generator"
node server.js
# Opens at http://localhost:3001
```

## Key Technical Details
- Chart.js 4.4.7 for visualizations (CDN)
- jsPDF 2.5.1 + autoTable 3.8.3 for PDF export (CDN)
- Express.js handles all server logic
- Data is cached in-memory and persisted to `kobo_data.xlsx`
- Server filters data on each API request (no pre-computed cache for filtered results)
- Dashboard HTML is ~3.8MB due to embedded JSON data
- `resolveUrl()` helper resolves relative redirect URLs against original domain
- `downloadFile()` has redirect depth limit (max 5) and HTTP status error handling

## Changelog
| Round | Changes |
|---|---|
| 1 | Core dashboard with metrics, charts, transaction table |
| 2 | Search, filters (Period, Month, Rep, Invoice), API data refresh |
| 3 | Top 20 Reps by Month filter, PDF export, Excel export |
| 4 | Multi-rep invoice handling (separators: / ; \) |
| 5 | Updated table display to show all reps per invoice |
| 6 | Top 20 ranked by invoice count, fixed Kobo API refresh redirect handling |
