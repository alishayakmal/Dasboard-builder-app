# Shay Analytics AI

A static, client-side web app that turns CSVs into an analysis-ready dashboard with KPIs, trends, and insights. Runs entirely in the browser and supports a lightweight signup gate with localStorage.

## Mock marketing feeds
The marketing page uses lightweight mock feeds located in `src/data/` and fetched via `src/lib/sampleFeed.js`. The data is rendered by DOM-based components in `src/components/marketing/`.

Upgrade path to React:
1. Keep the data contracts in `src/data/` (or replace with real API calls inside `getSampleInsights` and `getMiniDashboard`).
2. Replace `src/marketing.js` with a React entry point that imports the same feed helpers.
3. Port `renderSampleInsightsSection` and `renderMiniDashboardPreview` to React components without changing the data shapes.

## What it does
- Landing page with signup modal and localStorage gating
- Upload CSV and generate KPIs, charts, and tables
- Connect API (stub), Upload PDF (stub), Google Sheets import
- Sample datasets gallery
- Export filtered CSV and download insights

## Run locally
Windows users: use **Command Prompt** or **Git Bash** for npm commands (do not use PowerShell, which can block `npm.ps1`).

Backend (API + AI chat):
1. `cd backend`
1. `npm install`
1. Copy `.env.example` to `.env` and set `OPENAI_API_KEY`, `OPENAI_MODEL`, `PORT`, `FRONTEND_ORIGIN`.
1. `npm run dev`

Frontend (static + proxy):
1. From repo root: `npm install`
1. `npm run dev`
1. Open `http://localhost:3000`

Run both together (Windows-friendly):
1. `cd backend && npm install`
1. `cd .. && npm install`
1. `npm run dev:all`

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

## Chat Troubleshooting
- Set `OPENAI_API_KEY` in `backend/.env`.
- Create a key in **Project settings → API Keys → Create new secret key**.
- Reference: https://help.openai.com/en/articles/9186755-managing-your-work-in-the-api-platform-with-projects
- Run backend on port `5050` (or set `PORT`).
- The frontend dev server proxies `/api` to the backend.
- Common failure: `Unexpected token '<'` in chat means the proxy is misrouted or the backend is not running.

## Git Dependency Audit (SSH Error 128)
If you see `Permission denied (publickey)` during `npm install`, scan for forbidden SSH git deps:
1. `npm run audit:gitdeps`
2. If a lockfile is flagged:
   - Delete the lockfile(s)
   - Delete `node_modules`
   - Reinstall

## Local Validation Checklist
1. Use Command Prompt or Git Bash on Windows.
2. `cd backend` → `npm install` → create `backend/.env` with real key.
3. `npm run dev` and verify `http://localhost:5050/api/health` returns JSON.
4. From repo root: `npm install` then `npm run dev:all`.
5. In the browser, verify `/api/chat` responds with `Content-Type: application/json`.

## Leads Webhook Setup
Use a Google Apps Script Web App to capture signups in a sheet named **Leads**.

Deployment requirements:
- Deploy as **Web App**
- Execute as **Me**
- Who has access: **Anyone**
- Use the URL ending in `/exec`

In `app.js`, set:
`const WEBHOOK_URL = "YOUR_EXEC_URL_HERE";`
