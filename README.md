# Rupaiya 💰 — Personal Expense Tracker

**Free forever. Runs 100% locally. No login. No ads. No data sent anywhere.**

---

## Features

| Feature | Details |
|---|---|
| **Add Expenses** | Date, Amount (₹ INR), Description, Category, Nature, Paid By, Notes |
| **Import Excel** | Upload `.xlsx` / `.xls` / `.csv` with auto column mapping |
| **Export Excel** | Download all/filtered expenses as `.xlsx` |
| **Dashboard** | Summary cards + Trend, Category, Payment, Day-of-Week, Nature charts |
| **Analytics** | Heatmap, Insights, Top Expenses, Category vs Month, Budget Tracker |
| **Custom Charts** | Build your own charts — pick type, axis, measure, palette & save them |
| **Settings** | Manage categories, natures, payment methods; export/import JSON backup |
| **Dark Mode** | Full dark/light theme toggle |
| **Keyboard Shortcuts** | `N` = New expense, `Ctrl+I` = Import, `Ctrl+E` = Export, `Esc` = Close modal |

---

## Option 1 — Run in Browser (Zero install, free forever)

Just open `index.html` in any modern browser (Chrome, Edge, Firefox).

```
Double-click index.html   →   Opens in your browser   →   Done!
```

All data is saved in your **browser's localStorage** — it persists across sessions.

---

## Option 2 — Run as Desktop App (.exe) with Electron

### Prerequisites
- [Node.js](https://nodejs.org) (v18+)

### Run in development
```bash
npm install
npm start
```

### Build Windows .exe installer
```bash
npm install
npm run build-win
```
The installer will be in `dist/`. Run it once and Rupaiya is installed on your PC.

### Build for Mac
```bash
npm run build-mac
```

### Build for Linux
```bash
npm run build-linux
```

> **Electron is free and open-source.** electron-builder is also free. You pay nothing.

---

## Importing Excel Files

Your Excel file can have any column order. Rupaiya auto-detects columns named:

| Your Column Name | Maps To |
|---|---|
| `Date`, `Transaction Date`, `Txn Date` | Date |
| `Amount`, `Debit`, `Price`, `Rs`, `INR` | Amount |
| `Description`, `Narration`, `Particulars` | Description |
| `Category`, `Cat`, `Head`, `Type` | Category |
| `Nature`, `Kind`, `Classification` | Nature |
| `Paid By`, `Payment Method`, `Mode` | Paid By |
| `Notes`, `Remarks`, `Comments` | Notes |

You can also manually map any column to any field using the dropdown UI.

---

## Data Storage

- **Browser**: `localStorage` (survives browser restarts, cleared only if you clear site data)
- **Desktop (Electron)**: Same `localStorage` via Chromium engine
- Use **Settings → Export as JSON** to create a backup file anytime

---

## Tech Stack

- Pure HTML + CSS + JavaScript (no framework)
- [Chart.js](https://www.chartjs.org/) — charts
- [SheetJS (xlsx)](https://sheetjs.com/) — Excel import/export
- [Electron](https://www.electronjs.org/) — desktop packaging (optional)

All CDN libraries load from the internet on first use. For fully offline use, download the libs and reference them locally.
