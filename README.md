# Enterprise Demand Forecasting & Inventory Decision Support System

A static, browser-based dashboard for enterprise demand forecasting, inventory decision support, risk diagnostics, and financial analysis.

## Runtime files

- `index.html` - dashboard structure and content
- `styles.css` - visual styling and responsive behavior
- `app.js` - dashboard interactions, filtering, charts, and client-side logic
- `dashboard-data.json` - runtime data source loaded by the dashboard

## Run locally

Because the dashboard loads JSON with `fetch()`, open it through a local web server rather than `file://`.

### Python

```bash
python -m http.server 8000
```

Then open:

`http://localhost:8000`

## GitHub Pages

This repository includes a GitHub Actions workflow that deploys the root directory to GitHub Pages whenever changes are pushed to `main`.

In GitHub, go to **Settings -> Pages -> Build and deployment -> Source -> GitHub Actions**.

## Important data note

`dashboard-data.json` is required by the live dashboard and is therefore included in this repository package. The original Excel workbook is intentionally excluded from the GitHub-ready package because it is not required at runtime. Do not publish `dashboard-data.json` to a public repository if its contents are confidential or proprietary.

## External runtime dependencies

The dashboard currently loads Plotly and SheetJS from public CDNs, so an internet connection is required for those browser libraries.

## Backend-dependent endpoints

The frontend contains optional calls to `/report-results`, `/report-mobile`, and `/report-diag`. Those endpoints are not provided by this static GitHub Pages package; any functionality depending on them requires a separate backend service.
