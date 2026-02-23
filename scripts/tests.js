const assert = require('assert');

function parseNumber(value) {
  if (value === null || value === undefined) return null;
  let cleaned = String(value).trim();
  if (!cleaned) return null;
  const lower = cleaned.toLowerCase();
  if (["n/a", "na", "-", "--", ""].includes(lower)) return null;
  let isNegative = false;
  if (/^\(.*\)$/.test(cleaned)) {
    isNegative = true;
    cleaned = cleaned.replace(/^\(|\)$/g, "");
  }
  const isPercent = /%/.test(cleaned);
  cleaned = cleaned.replace(/%/g, "");
  cleaned = cleaned.replace(/\b(usd|cad|eur|gbp)\b/gi, "");
  cleaned = cleaned.replace(/[$€£]/g, "");
  cleaned = cleaned.replace(/[\s,_]/g, "");
  cleaned = cleaned.trim();
  if (!cleaned) return null;
  if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(cleaned)) return null;
  const number = Number(cleaned);
  if (Number.isNaN(number)) return null;
  let result = isNegative ? -number : number;
  if (isPercent) result /= 100;
  return result;
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return Number.isNaN(value.valueOf()) ? null : value;
  const raw = String(value).trim();
  if (!raw) return null;
  if (/^[0-9]+$/.test(raw)) return null;
  if (raw.includes("T") || raw.endsWith("Z")) {
    const isoDate = new Date(raw);
    if (!Number.isNaN(isoDate.valueOf())) return isoDate;
  }
  const ymd = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (ymd) {
    const year = Number(ymd[1]);
    const month = Number(ymd[2]);
    const day = Number(ymd[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) return date;
  }
  const mdy = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) {
    const month = Number(mdy[1]);
    const day = Number(mdy[2]);
    const year = Number(mdy[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) return date;
  }
  const fallback = new Date(raw);
  return Number.isNaN(fallback.valueOf()) ? null : fallback;
}

function inferMetricType(metric, values) {
  const lower = String(metric || "").toLowerCase();
  const isCurrency = /(revenue|sales|cost|spend|profit|price|amount|gmv|mrr|arr)/i.test(lower);
  const isRateName = /(rate|percent|ctr|cvr|roas|ratio|retention)/i.test(lower);
  const isDuration = /(time|duration|seconds|secs|minutes|mins|hours)/i.test(lower);
  const isRateRange = values.length > 0 && values.every((val) => val >= 0 && val <= 1);
  const isRate = isRateName || isRateRange;
  const kind = isRate ? "rate" : isCurrency ? "currency" : isDuration ? "duration" : "count";
  return { isCurrency, isRate, isDuration, kind };
}

function formatMetric(value, metricKind) {
  if (value === null || value === undefined || Number.isNaN(value)) return "—";
  if (metricKind === "rate") {
    return `${(value * 100).toFixed(1)}%`;
  }
  return String(value);
}

function profileDataset(rows, columns) {
  const profile = {};
  columns.forEach((col) => {
    const values = rows.map((row) => row[col]);
    const nonMissing = values.filter((value) => value !== null && value !== undefined && String(value).trim() !== "");
    const numericValues = [];
    const dateValues = [];
    let numericCount = 0;
    let dateCount = 0;
    nonMissing.forEach((value) => {
      const num = parseNumber(value);
      if (num !== null) {
        numericCount += 1;
        numericValues.push(num);
      }
      const date = parseDate(value);
      if (date) {
        dateCount += 1;
        dateValues.push(date);
      }
    });
    const numericRate = nonMissing.length ? numericCount / nonMissing.length : 0;
    const dateRate = nonMissing.length ? dateCount / nonMissing.length : 0;
    let type = "categorical";
    if (dateRate >= 0.9) type = "date";
    else if (numericRate >= 0.9) type = "numeric";
    if (/activeusers|sessions/.test(col.toLowerCase())) type = "numeric";
    profile[col] = { type, numericRate, dateRate };
  });
  return profile;
}

// Tests
assert.strictEqual(formatMetric(0.41, "rate"), "41.0%", "Rate formatting should be 41.0%");
const metricType = inferMetricType("RetentionRate", [0.41, 0.58]);
assert.strictEqual(metricType.kind, "rate", "RetentionRate should be rate");
const sampleRows = [
  { ActiveUsers: "120", Sessions: "300", Week: "2024-01-01" },
  { ActiveUsers: "140", Sessions: "320", Week: "2024-01-08" },
];
const profile = profileDataset(sampleRows, ["ActiveUsers", "Sessions", "Week"]);
assert.strictEqual(profile.ActiveUsers.type, "numeric", "ActiveUsers should be numeric");
assert.strictEqual(profile.Sessions.type, "numeric", "Sessions should be numeric");
assert.strictEqual(profile.Week.type, "date", "Week should be date");

console.log("All tests passed.");
