const path = require("path");
const express = require("express");
const cookieParser = require("cookie-parser");
const crypto = require("crypto");
const { randomState, buildAuthUrl, exchangeCode } = require("./googleAuth");
const { setSession, getSession, setState, consumeState } = require("./tokenStore");

const app = express();

const PORT = process.env.PORT || 3000;
const ROOT_DIR = path.resolve(__dirname, "..");
const SESSION_SECRET = process.env.SESSION_SECRET || "dev-secret-change-me";
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI;

if (!CLIENT_ID || !CLIENT_SECRET || !REDIRECT_URI) {
  console.warn("Missing Google OAuth env vars. OAuth routes will error until configured.");
}

app.use(express.json());
app.use(cookieParser(SESSION_SECRET));
app.use(express.static(ROOT_DIR));

app.get("/oauth/start", async (req, res) => {
  if (!CLIENT_ID || !CLIENT_SECRET || !REDIRECT_URI) {
    return res.status(500).json({ ok: false, error: "Missing Google OAuth environment variables." });
  }
  const state = randomState();
  await setState(state);
  const url = buildAuthUrl({ clientId: CLIENT_ID, redirectUri: REDIRECT_URI, state });
  res.redirect(url);
});

app.get("/oauth/callback", async (req, res) => {
  const { code, state, error } = req.query;
  if (error) {
    return res.status(400).json({ ok: false, error });
  }
  if (!code || !state) {
    return res.status(400).json({ ok: false, error: "Missing code or state." });
  }
  const validState = await consumeState(state);
  if (!validState) {
    return res.status(400).json({ ok: false, error: "Invalid or expired state." });
  }
  try {
    const tokens = await exchangeCode({
      code,
      clientId: CLIENT_ID,
      clientSecret: CLIENT_SECRET,
      redirectUri: REDIRECT_URI,
    });
    const sessionId = crypto.randomBytes(16).toString("hex");
    await setSession(sessionId, tokens);
    res.cookie("saai_session", sessionId, {
      httpOnly: true,
      secure: true,
      sameSite: "lax",
      maxAge: 60 * 60 * 1000,
    });
    res.redirect("/#/app?google=connected");
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message });
  }
});

app.get("/api/sheets/read", async (req, res) => {
  const sessionId = req.cookies.saai_session;
  if (!sessionId) {
    return res.status(401).json({ ok: false, error: "Not authenticated." });
  }
  const tokens = await getSession(sessionId);
  if (!tokens?.access_token) {
    return res.status(401).json({ ok: false, error: "Session expired." });
  }

  const { spreadsheetId, range } = req.query;
  if (!spreadsheetId || !range) {
    return res.status(400).json({ ok: false, error: "Missing spreadsheetId or range." });
  }

  try {
    const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;
    const response = await fetch(apiUrl, {
      headers: { Authorization: `Bearer ${tokens.access_token}` },
    });
    const bodyText = await response.text();
    if (!response.ok) {
      return res.status(response.status).json({ ok: false, error: bodyText });
    }
    return res.status(200).json(JSON.parse(bodyText));
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("*", (req, res) => {
  res.sendFile(path.join(ROOT_DIR, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
