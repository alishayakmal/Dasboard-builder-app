<<<<<<< HEAD
﻿# Data to Dashboard

A static, client-side web app that turns any CSV into an analysis-driven dashboard with KPI selection, recommended charts, a correlation heat map, and insight statements.

## What it does
- Profiles the dataset on every upload
- Validates data quality and flags warnings
- Auto-detects numeric, categorical, and date columns
- Selects meaningful KPIs and renders summary stats
- Generates trend, breakdown, and analysis charts
- Includes a correlation heat map and insights panel
- Provides interactive controls for metric, dimension, date range, and Top N

## Run locally
Open `index.html` directly in your browser. No build tools required.

## Usage
1. Click **Upload CSV** or **Load Sample**.
2. Review KPI tiles and warnings.
3. Adjust the primary metric, dimension, date range, and Top N controls.
4. Use insights and charts to guide next steps.

## Sample data
The included sample file is `data-sample.csv`.

## Deploy on GitHub Pages
1. Push the repo to GitHub.
2. In GitHub, go to **Settings → Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Choose **main** branch and **/ (root)**.
5. Click **Save**.

Your live site will be available at:
`https://YOUR_USERNAME.github.io/Dasboard-builder-app/`
=======
﻿# Data to Dashboard

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

>>>>>>> 2dc1070076742e75405e5b15c993a338e3cc61aa
