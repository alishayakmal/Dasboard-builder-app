# Shay Analytics AI Backend

Express backend to handle Google OAuth and private Google Sheets access.

## Setup
1. `cd backend`
2. `npm install`
3. Copy `.env.example` to `.env` and fill in values:
   - `GOOGLE_CLIENT_ID`
   - `GOOGLE_CLIENT_SECRET`
   - `GOOGLE_REDIRECT_URI` (e.g. `http://localhost:8787/auth/google/callback`)
   - `SESSION_SECRET`
   - `SPREADSHEET_ID`
   - `SHEET_NAME`
   - `FRONTEND_ORIGIN` (e.g. `http://localhost:8000` or your GitHub Pages URL)
4. `npm run dev`

## Routes
- `GET /health` → `{ ok: true }`
- `GET /auth/google` → starts OAuth
- `GET /auth/google/callback` → OAuth callback
- `POST /api/sheets/append` → append a row
- `GET /api/sheets/read` → read sheet values
