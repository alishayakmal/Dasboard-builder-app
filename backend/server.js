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

function normalizeMessage(message) {
  return String(message || "").toLowerCase();
}

function resolveFieldFromMessage(message, schema, fallbackKeywords = []) {
  const columns = Array.isArray(schema?.columns) ? schema.columns.map((col) => col?.name).filter(Boolean) : [];
  const lower = normalizeMessage(message);
  let match = columns.find((name) => lower.includes(name.toLowerCase()));
  if (match) return match;

  const tokens = lower.split(/[\s,]+/).filter(Boolean);
  let best = null;
  let bestScore = 0;
  columns.forEach((name) => {
    const nameLower = name.toLowerCase();
    let score = 0;
    tokens.forEach((token) => {
      if (token.length < 3) return;
      if (nameLower.includes(token)) score += 1;
    });
    fallbackKeywords.forEach((token) => {
      if (nameLower.includes(token)) score += 2;
    });
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  });
  return best;
}

function parseTopN(message) {
  const match = message.match(/top\s+(\d+)/i);
  if (!match) return null;
  const n = Number(match[1]);
  return Number.isFinite(n) ? n : null;
}

function parseAggregation(message) {
  const lower = normalizeMessage(message);
  if (lower.includes("average") || lower.includes("avg")) return "avg";
  if (lower.includes("sum")) return "sum";
  if (lower.includes("count distinct")) return "count_distinct";
  if (lower.includes("count")) return "count";
  return null;
}

function parseChartType(message) {
  const lower = normalizeMessage(message);
  if (lower.includes("bar chart") || lower.includes("bar")) return "bar";
  if (lower.includes("line chart") || lower.includes("line")) return "line";
  return null;
}

function parseFilter(message) {
  const match = message.match(/filter\s+(.+?)\s+to\s+(.+)/i);
  if (!match) return null;
  return { fieldHint: match[1], value: match[2] };
}

function applyToolCallsToState(state, toolCalls) {
  let next = { ...state };
  toolCalls.forEach((call) => {
    const args = call.args || {};
    if (call.name === "setMetric") next.selectedMetric = args.field;
    if (call.name === "setDimension") next.selectedDimension = args.field;
    if (call.name === "setChartType") next.chartType = args.type;
    if (call.name === "setAggregation") next.aggregation = args.agg;
    if (call.name === "setTopN") next.topN = args.n;
    if (call.name === "setSort") next.sort = { field: args.field, direction: args.direction };
    if (call.name === "applyFilter") {
      const filters = Array.isArray(next.filters) ? [...next.filters] : [];
      const idx = filters.findIndex((f) => f.field === args.field);
      const nextFilter = { field: args.field, operator: args.operator, value: args.value };
      if (idx >= 0) filters[idx] = nextFilter;
      else filters.push(nextFilter);
      next.filters = filters;
    }
    if (call.name === "clearFilter") {
      next.filters = (next.filters || []).filter((f) => f.field !== args.field);
    }
  });
  return next;
}

app.post("/api/chat", async (req, res) => {
  try {
    const { message, datasetSchema, currentDashboardState } = req.body || {};
    if (!message) {
      return res.status(400).json({ ok: false, error: "Missing message." });
    }
    const toolCalls = [];
    const chartType = parseChartType(message);
    if (chartType) {
      toolCalls.push({ name: "setChartType", args: { type: chartType } });
    }

    const aggregation = parseAggregation(message);
    if (aggregation) {
      toolCalls.push({ name: "setAggregation", args: { agg: aggregation } });
    }

    const topN = parseTopN(message);
    if (topN) {
      toolCalls.push({ name: "setTopN", args: { n: topN } });
    }

    const byMatch = message.match(/by\s+([^.;]+)/i);
    if (byMatch) {
      const byHint = byMatch[1];
      const field = resolveFieldFromMessage(byHint, datasetSchema, ["org", "name", "company"]);
      if (field) toolCalls.push({ name: "setDimension", args: { field } });
    }

    if (/metric|by\s+/i.test(message)) {
      const metricField = resolveFieldFromMessage(message, datasetSchema, ["count", "unique", "total"]);
      if (metricField && /top\s+\d+/i.test(message)) {
        toolCalls.push({ name: "setMetric", args: { field: metricField } });
      }
    }

    const filter = parseFilter(message);
    if (filter) {
      const field = resolveFieldFromMessage(filter.fieldHint, datasetSchema, ["org", "name", "company"]);
      if (field) {
        toolCalls.push({ name: "applyFilter", args: { field, operator: "==", value: filter.value } });
      }
    }

    if (!toolCalls.length) {
      return res.status(200).json({
        ok: false,
        error: "I could not map that request to a supported chart action. Try rephrasing.",
      });
    }

    const nextState = applyToolCallsToState(currentDashboardState || {}, toolCalls);
    const proposedDiff = {
      before: currentDashboardState || {},
      after: nextState,
    };

    return res.json({
      ok: true,
      toolCalls,
      assistantMessage: "Here is a preview of the requested chart updates.",
      proposedDiff,
      notes: "Actions are returned as validated tool calls.",
    });
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
});

const port = process.env.PORT || 8787;
app.listen(port, () => {
  console.log(`Backend listening on ${port}`);
});
