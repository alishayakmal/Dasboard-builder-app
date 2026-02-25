const dateTokens = [
  "date",
  "time",
  "timestamp",
  "datetime",
  "created_at",
  "updated_at",
  "run_date",
  "day",
  "month",
  "year",
];

function nameLooksTemporal(name) {
  const key = String(name || "").toLowerCase().replace(/\s+/g, "_");
  return dateTokens.some((token) => key.includes(token));
}

export function detectDateField({ columns = [], profiles = {}, rows = [] } = {}) {
  const candidates = Object.entries(profiles || {})
    .filter(([_, profile]) => profile?.type === "date")
    .sort((a, b) => (b[1].dateRate || 0) - (a[1].dateRate || 0));
  if (candidates.length) {
    const [field, profile] = candidates[0];
    return { field, confidence: profile?.dateRate || 0.6 };
  }

  const fallback = (columns || []).find((col) => nameLooksTemporal(col));
  if (fallback) return { field: fallback, confidence: 0.3 };

  if (rows.length && columns.length) {
    const sample = rows.slice(0, 80);
    let best = null;
    let bestRate = 0;
    columns.forEach((col) => {
      let parsed = 0;
      let total = 0;
      sample.forEach((row) => {
        const value = row?.[col];
        if (value == null || value === "") return;
        total += 1;
        const date = new Date(value);
        if (!Number.isNaN(date.getTime())) parsed += 1;
      });
      const rate = total ? parsed / total : 0;
      if (rate > bestRate) {
        bestRate = rate;
        best = col;
      }
    });
    if (best && bestRate >= 0.6) return { field: best, confidence: bestRate };
  }

  return { field: null, confidence: 0 };
}
