# Data to Dashboard

A static, client-side web app that turns any CSV into a clean dashboard with KPIs, charts, and a sortable data table. Runs entirely in the browser.

## What it does
- Upload a CSV via file picker
- Auto-detect numeric columns as metrics
- Detect a date column for time-series trend charts
- Render KPI tiles (3–6 metrics)
- Render a line chart and bar chart
- Render a sortable data table preview

## Run locally
Open `index.html` directly in your browser.

## Deploy on GitHub Pages
1. Push the repo to GitHub.
2. In GitHub, go to **Settings → Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Choose **main** branch and **/ (root)**.
5. Click **Save**.

Your live site will be available at:
 https://alishayakmal.github.io/Dasboard-builder-app/

## Live demo
Add your GitHub Pages URL here once enabled.

## Notes
All processing is done client-side using PapaParse and Chart.js loaded via CDN.

