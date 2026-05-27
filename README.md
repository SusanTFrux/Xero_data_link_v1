# Xero Data Link — Excel Add-in v1.0

Pull live Profit & Loss and Balance Sheet data from Xero directly into Excel. Appears as a button in the Excel ribbon, opens a sidebar, and writes data straight into your cells.

## Features (v1)

- **P&L and Balance Sheet** — select either report type
- **Live account list** — accounts loaded from your Xero chart of accounts, grouped by type
- **Custom date range** — monthly columns, any start/end date
- **Tracking category filter** — filter by project, department, or any Xero tracking category
- **Grouped output** — section headers, subtotals with SUM formulas, and Net Profit / Net Assets grand total
- **One-click refresh** — saved link config stored in the workbook; re-fetch with one click
- **Auto-extend dates** — detects when a new month is available and offers to include it
- **Fluent UI design** — Microsoft's design system; looks native to Office

## Files

| File | Purpose |
|---|---|
| `taskpane.html` | Sidebar UI |
| `taskpane.js` | All logic: OAuth, Excel read/write, Xero API, wizard |
| `styles.css` | Fluent UI 2 stylesheet |
| `auth-dialog.html` | OAuth popup for Xero login |
| `manifest.xml` | Registers the add-in with Excel |
| `netlify/functions/token.js` | Serverless function: Xero token exchange |
| `netlify.toml` | Netlify build config |
| `assets/` | Ribbon icons |

## Deployment

### 1. Upload to GitHub
- Create a new repository at github.com
- Upload all files from this folder (not the folder itself — the contents)

### 2. Deploy to Netlify
- Go to netlify.com → Add new site → Import from GitHub
- Select your repository → Deploy with defaults
- Copy your Netlify URL (e.g. `amazing-fox-123.netlify.app`)

### 3. Add Xero secrets to Netlify
- Site configuration → Environment variables → Add:
  - `XERO_CLIENT_ID` = your Client ID
  - `XERO_CLIENT_SECRET` = your Client Secret
- Trigger a redeploy

### 4. Update Xero redirect URI
- developer.xero.com → your app → Configuration
- Add redirect URI: `https://YOUR-APP.netlify.app/auth-dialog.html`

### 5. Sideload into Excel
- Insert → Add-ins → My Add-ins → Upload My Add-in → browse to `manifest.xml`

## Required Xero Scopes

- `accounting.reports.profitandloss.read`
- `accounting.reports.balancesheet.read`
- `accounting.settings.read`
- `offline_access`

## v2 Roadmap

- Multiple data links per workbook
- Financial Year and YTD column options
- Quarterly aggregation
