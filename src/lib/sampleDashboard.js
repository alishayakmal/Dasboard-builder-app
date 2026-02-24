const SAMPLE_DATA_PATH = "./src/sample_data/sample_insights.csv";

const METRICS = [
  {
    key: "revenue",
    label: "Revenue",
    unit: "currency",
    definition: "Sum of Revenue over the selected period.",
  },
  {
    key: "retention",
    label: "Retention",
    unit: "rate",
    definition: "Retained users / eligible users.",
  },
  {
    key: "activeUsers",
    label: "Active users",
    unit: "count",
    definition: "Sum of active users over the selected period.",
  },
  {
    key: "pipeline",
    label: "Pipeline",
    unit: "currency",
    definition: "Sum of pipeline amount over the selected period.",
  },
];

const formatter = new Intl.NumberFormat("en-US");

function formatValue(value, unit) {
  if (unit === "currency") {
    return `$${formatter.format(Math.round(value))}`;
  }
  if (unit === "rate") {
    return `${(value * 100).toFixed(1)}%`;
  }
  return formatter.format(Math.round(value));
}

function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines[0].split(",");
  return lines.slice(1).map((line) => {
    const values = line.split(",");
    const row = {};
    headers.forEach((header, index) => {
      row[header] = values[index];
    });
    return row;
  });
}

function toNumber(value) {
  const num = Number(value);
  return Number.isNaN(num) ? 0 : num;
}

function toDate(value) {
  return new Date(value);
}

export async function loadSampleDataset() {
  const response = await fetch(SAMPLE_DATA_PATH);
  if (!response.ok) throw new Error("Sample dataset unavailable");
  const text = await response.text();
  return parseCsv(text).map((row) => ({
    date: toDate(row.date),
    plan: row.plan,
    channel: row.channel,
    region: row.region,
    revenue: toNumber(row.revenue),
    activeUsers: toNumber(row.activeUsers),
    retainedUsers: toNumber(row.retainedUsers),
    eligibleUsers: toNumber(row.eligibleUsers),
    pipelineAmount: toNumber(row.pipelineAmount),
  }));
}

export function getMetricConfig() {
  return METRICS;
}

export function computePeriods(rows, windowDays = 30) {
  const dates = rows.map((row) => row.date).sort((a, b) => a - b);
  const end = dates[dates.length - 1];
  const start = new Date(end);
  start.setDate(start.getDate() - (windowDays - 1));
  const priorEnd = new Date(start);
  priorEnd.setDate(priorEnd.getDate() - 1);
  const priorStart = new Date(priorEnd);
  priorStart.setDate(priorStart.getDate() - (windowDays - 1));
  return { start, end, priorStart, priorEnd };
}

export function computeKpis(rows, windowDays = 30) {
  const { start, end, priorStart, priorEnd } = computePeriods(rows, windowDays);
  const current = rows.filter((row) => row.date >= start && row.date <= end);
  const prior = rows.filter((row) => row.date >= priorStart && row.date <= priorEnd);

  const sum = (items, key) => items.reduce((acc, row) => acc + row[key], 0);
  const sumPipeline = (items) => items.reduce((acc, row) => acc + row.pipelineAmount, 0);
  const retention = (items) => {
    const retained = sum(items, "retainedUsers");
    const eligible = sum(items, "eligibleUsers");
    return eligible ? retained / eligible : 0;
  };

  const revenueCurrent = sum(current, "revenue");
  const revenuePrior = sum(prior, "revenue");
  const activeCurrent = sum(current, "activeUsers");
  const activePrior = sum(prior, "activeUsers");
  const pipelineCurrent = sumPipeline(current);
  const pipelinePrior = sumPipeline(prior);
  const retentionCurrent = retention(current);
  const retentionPrior = retention(prior);

  return [
    {
      key: "revenue",
      label: "Revenue",
      value: revenueCurrent,
      delta: revenuePrior ? (revenueCurrent - revenuePrior) / revenuePrior : 0,
      deltaType: "percent",
      window: "Last 30 days vs prior 30 days",
      definition: METRICS.find((m) => m.key === "revenue").definition,
    },
    {
      key: "retention",
      label: "Retention",
      value: retentionCurrent,
      delta: retentionCurrent - retentionPrior,
      deltaType: "percentagePoints",
      window: "Last 30 days vs prior 30 days",
      definition: METRICS.find((m) => m.key === "retention").definition,
    },
    {
      key: "activeUsers",
      label: "Active users",
      value: activeCurrent,
      delta: activePrior ? (activeCurrent - activePrior) / activePrior : 0,
      deltaType: "percent",
      window: "Last 30 days vs prior 30 days",
      definition: METRICS.find((m) => m.key === "activeUsers").definition,
    },
    {
      key: "pipeline",
      label: "Pipeline",
      value: pipelineCurrent,
      delta: pipelinePrior ? (pipelineCurrent - pipelinePrior) / pipelinePrior : 0,
      deltaType: "percent",
      window: "Last 30 days vs prior 30 days",
      definition: METRICS.find((m) => m.key === "pipeline").definition,
    },
  ];
}

export function computeTrend(rows, metricKey, breakdown) {
  const groupKey = breakdown === "overall" ? null : breakdown;
  const groups = {};
  rows.forEach((row) => {
    const month = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, "0")}`;
    const group = groupKey ? row[groupKey] : "Overall";
    if (!groups[group]) groups[group] = {};
    if (!groups[group][month]) groups[group][month] = [];
    groups[group][month].push(row);
  });

  const series = Object.entries(groups).map(([group, months]) => {
    const points = Object.entries(months)
      .map(([month, items]) => {
        const value = computeMetricValue(items, metricKey);
        return { month, value };
      })
      .sort((a, b) => a.month.localeCompare(b.month));
    return { group, points };
  });

  return series;
}

export function computeBreakdown(rows, metricKey, dimension, windowDays = 30) {
  const { start, end } = computePeriods(rows, windowDays);
  const current = rows.filter((row) => row.date >= start && row.date <= end);
  const buckets = {};
  current.forEach((row) => {
    const key = row[dimension];
    if (!buckets[key]) buckets[key] = [];
    buckets[key].push(row);
  });
  return Object.entries(buckets).map(([key, items]) => ({
    key,
    value: computeMetricValue(items, metricKey),
    sampleSize: items.length,
  })).sort((a, b) => b.value - a.value);
}

export function computeInsights(rows, metricKey, dimension) {
  const breakdown = computeBreakdown(rows, metricKey, dimension);
  if (!breakdown.length) return [];
  const top = breakdown[0];
  const bottom = breakdown[breakdown.length - 1];
  const metricLabel = METRICS.find((m) => m.key === metricKey)?.label || metricKey;
  const delta = top.value - bottom.value;
  const deltaLabel = metricKey === "retention" ? `${(delta * 100).toFixed(1)} pp` : formatValue(delta, METRICS.find((m) => m.key === metricKey)?.unit);

  return [
    {
      claim: `${metricLabel} is higher in ${top.key} than ${bottom.key}.`,
      evidence: `Δ ${deltaLabel} (${top.key} vs ${bottom.key}) · n=${top.sampleSize}/${bottom.sampleSize}`,
      driver: `Top segment: ${top.key} (${formatValue(top.value, METRICS.find((m) => m.key === metricKey)?.unit)})`,
      action: `Review ${dimension} drivers to replicate ${top.key} performance.`,
      confidence: top.sampleSize >= 5 && bottom.sampleSize >= 5 ? "medium" : "low",
    },
  ];
}

export { formatValue };
