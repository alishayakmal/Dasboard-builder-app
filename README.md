# Shalytics AI

A static, client-side web app that turns CSVs into an analysis-ready dashboard with KPIs, trends, and insights. Runs entirely in the browser and supports a lightweight signup gate with localStorage.

## What it does
- Landing page with signup modal and localStorage gating
- Upload CSV and generate KPIs, charts, and tables
- Connect API (stub), Upload PDF (stub), Google Sheets import
- Sample datasets gallery
- Export filtered CSV and download insights

## Run locally
Open `index.html` directly in your browser. No build tools required.

## Deploy on GitHub Pages
1. Push the repo to GitHub.
2. In GitHub, go to **Settings → Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Choose **main** branch and **/ (root)**.
5. Click **Save**.

Your live site will be available at:
`https://YOUR_USERNAME.github.io/Dasboard-builder-app/`

## Private Sheets Setup
This project includes a small Express backend to authenticate Google Sheets via OAuth and read private sheets.

Environment variables:
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REDIRECT_URI` (e.g., `http://localhost:3000/oauth/callback`)
- `SESSION_SECRET`

Flow:
1. User clicks **Connect Google** (calls `/oauth/start`).
2. Google redirects to `/oauth/callback`, which stores tokens server-side and sets an HttpOnly session cookie.
3. The app calls `/api/sheets/read` with `spreadsheetId` and `range` to fetch values.

Manual test plan:
1. Run `npm install`, then `npm run dev`.
2. Click **Connect Google** and approve access.
3. Paste a private Google Sheets link and load it; dashboard should render.
4. If you use a public sheet, it should also load without auth.

## Leads Webhook Setup
Use a Google Apps Script Web App to capture signups in a sheet named **Leads**.

Deployment requirements:
- Deploy as **Web App**
- Execute as **Me**
- Who has access: **Anyone**
- Use the URL ending in `/exec`

In `app.js`, set:
`const WEBHOOK_URL = "YOUR_EXEC_URL_HERE";`
