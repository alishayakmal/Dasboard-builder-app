const express = require("express");
const session = require("express-session");
const cookieParser = require("cookie-parser");
const cors = require("cors");
const crypto = require("crypto");
const { google } = require("googleapis");
const { setRefreshToken, getRefreshToken, setState, consumeState } = require("./db");

const {
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  GOOGLE_REDIRECT_URI,
  SESSION_SECRET,
  SPREADSHEET_ID,
  SHEET_NAME,
  FRONTEND_ORIGIN,
} = process.env;

const app = express();

app.use(express.json());
app.use(cookieParser());
app.use(
  cors({
    origin: FRONTEND_ORIGIN,
    credentials: true,
  })
);
app.use(
  session({
    secret: SESSION_SECRET || "dev-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
    },
  })
);

const oauth2Client = new google.auth.OAuth2(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REDIRECT_URI);
const sheets = google.sheets({ version: "v4", auth: oauth2Client });

function requireAuth(req, res, next) {
  if (!req.session || !req.session.id) {
    return res.status(401).json({ ok: false, error: "Not authenticated" });
  }
  next();
}

app.get("/health", (req, res) => {
  res.json({ ok: true });
});

app.get("/auth/google", async (req, res) => {
  const state = crypto.randomBytes(16).toString("hex");
  await setState(state);
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: ["https://www.googleapis.com/auth/spreadsheets"],
    state,
  });
  res.redirect(url);
});

app.get("/auth/google/callback", async (req, res) => {
  const { code, state } = req.query;
  if (!code || !state) {
    return res.status(400).json({ ok: false, error: "Missing code or state" });
  }
  const validState = await consumeState(state);
  if (!validState) {
    return res.status(400).json({ ok: false, error: "Invalid state" });
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    if (tokens.refresh_token) {
      await setRefreshToken(req.session.id, tokens.refresh_token);
    }
    res.redirect(FRONTEND_ORIGIN);
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post("/api/sheets/append", requireAuth, async (req, res) => {
  try {
    const refreshToken = await getRefreshToken(req.session.id);
    if (!refreshToken) {
      return res.status(401).json({ ok: false, error: "Missing refresh token" });
    }
    oauth2Client.setCredentials({ refresh_token: refreshToken });
    const payload = req.body || {};
    const row = [
      new Date().toISOString(),
      payload.source || "",
      payload.user || "",
      payload.action || "",
      payload.entity || "",
      payload.value || "",
      JSON.stringify(payload),
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:G`,
      valueInputOption: "RAW",
      requestBody: { values: [row] },
    });

    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message });
  }
});

app.get("/api/sheets/read", requireAuth, async (req, res) => {
  try {
    const refreshToken = await getRefreshToken(req.session.id);
    if (!refreshToken) {
      return res.status(401).json({ ok: false, error: "Missing refresh token" });
    }
    oauth2Client.setCredentials({ refresh_token: refreshToken });
    const range = req.query.range || `${SHEET_NAME}!A:Z`;
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
    });
    res.json(result.data);
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message });
  }
});

const port = process.env.PORT || 8787;
app.listen(port, () => {
  console.log(`Backend listening on ${port}`);
});
