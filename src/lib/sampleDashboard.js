const SAMPLE_DATA_PATH = "./src/sample_data/sample_insights.csv";
const FALLBACK_CSV = `date,plan,channel,region,revenue,activeUsers,retainedUsers,eligibleUsers,pipelineAmount
2024-01-03,Enterprise,Paid Search,NA,42000,1200,520,1000,78000
2024-01-10,Enterprise,Organic,NA,43800,1260,530,1020,82000
2024-01-17,Enterprise,Paid Social,EU,44500,1310,545,1040,83500
2024-01-24,Enterprise,Paid Search,EU,46200,1360,560,1060,86000
2024-01-31,Enterprise,Organic,APAC,47800,1400,575,1080,88000
2024-02-07,Enterprise,Paid Search,NA,51200,1500,610,1110,94000
2024-02-14,Enterprise,Paid Social,EU,52600,1540,625,1130,96000
2024-02-21,Enterprise,Organic,APAC,54800,1600,640,1160,99000
2024-02-28,Enterprise,Paid Search,NA,57200,1660,660,1190,103000
2024-03-06,Enterprise,Organic,EU,59800,1720,680,1220,107000
2024-03-13,Enterprise,Paid Social,APAC,61200,1780,698,1250,110000
2024-03-20,Enterprise,Paid Search,NA,64000,1860,720,1280,115000
2024-01-03,Mid-market,Paid Search,NA,26000,980,410,920,52000
2024-01-10,Mid-market,Organic,NA,24800,960,402,910,50500
2024-01-17,Mid-market,Paid Social,EU,25200,970,405,915,51200
2024-01-24,Mid-market,Paid Search,EU,25800,980,412,920,52000
2024-01-31,Mid-market,Organic,APAC,26400,990,418,925,52800
2024-02-07,Mid-market,Paid Search,NA,27200,1010,424,935,54000
2024-02-14,Mid-market,Paid Social,EU,26800,1000,420,930,53500
2024-02-21,Mid-market,Organic,APAC,27500,1020,426,940,54800
2024-02-28,Mid-market,Paid Search,NA,28100,1030,432,950,56000
2024-03-06,Mid-market,Organic,EU,28600,1040,438,960,57000
2024-03-13,Mid-market,Paid Social,APAC,29200,1060,444,970,58500
2024-03-20,Mid-market,Paid Search,NA,29800,1070,450,980,59500
2024-01-03,SMB,Organic,NA,14000,720,280,700,30000
2024-01-10,SMB,Paid Social,NA,14500,735,286,710,31000
2024-01-17,SMB,Paid Search,EU,14800,745,290,720,31500
2024-01-24,SMB,Organic,EU,15100,755,294,730,32000
2024-01-31,SMB,Paid Social,APAC,15400,765,298,740,32500
2024-02-07,SMB,Paid Search,NA,15800,780,304,755,33500
2024-02-14,SMB,Organic,EU,16200,795,310,770,34500
2024-02-21,SMB,Paid Social,APAC,16600,810,316,785,35500
2024-02-28,SMB,Paid Search,NA,17000,825,322,800,36500
2024-03-06,SMB,Organic,EU,17500,840,328,815,37500
2024-03-13,SMB,Paid Social,APAC,18000,860,334,830,39000
2024-03-20,SMB,Paid Search,NA,18600,880,342,850,41000`;

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
  let text = "";
  try {
    const response = await fetch(SAMPLE_DATA_PATH);
    if (!response.ok) throw new Error("Sample dataset unavailable");
    text = await response.text();
  } catch (error) {
    text = FALLBACK_CSV;
  }
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
    const monthKey = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, "0")}-01`;
    const group = groupKey ? row[groupKey] : "Overall";
    if (!groups[group]) groups[group] = {};
    if (!groups[group][monthKey]) groups[group][monthKey] = [];
    groups[group][monthKey].push(row);
  });

  const allMonthKeys = new Set();
  Object.values(groups).forEach((months) => {
    Object.keys(months).forEach((key) => allMonthKeys.add(key));
  });
  const sortedMonths = fillMissingMonths(Array.from(allMonthKeys).sort());

  return Object.entries(groups).map(([group, months]) => {
    const points = sortedMonths.map((month) => {
      const items = months[month] || [];
      const value = items.length ? computeMetricValue(items, metricKey) : 0;
      return { month, monthTs: new Date(month).getTime(), value };
    });
    return { group, points };
  });
}

function fillMissingMonths(sortedMonthKeys) {
  if (!sortedMonthKeys.length) return [];
  const start = new Date(sortedMonthKeys[0]);
  const end = new Date(sortedMonthKeys[sortedMonthKeys.length - 1]);
  const keys = [];
  const cursor = new Date(start);
  while (cursor <= end) {
    const key = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}-01`;
    keys.push(key);
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return keys;
}

function computeMetricValue(items, metricKey) {
  const sum = (key) => items.reduce((acc, row) => acc + row[key], 0);
  if (metricKey === "revenue") return sum("revenue");
  if (metricKey === "pipeline") return sum("pipelineAmount");
  if (metricKey === "activeUsers") return sum("activeUsers");
  if (metricKey === "retention") {
    const retained = sum("retainedUsers");
    const eligible = sum("eligibleUsers");
    return eligible ? retained / eligible : 0;
  }
  return 0;
}

function getMetricFieldKeys(metricKey) {
  if (metricKey === "retention") return ["retainedUsers", "eligibleUsers"];
  if (metricKey === "pipeline") return ["pipelineAmount"];
  if (metricKey === "activeUsers") return ["activeUsers"];
  if (metricKey === "revenue") return ["revenue"];
  return [];
}

function buildMonthlyMetricMap(items, metricKey) {
  const buckets = {};
  items.forEach((row) => {
    const monthKey = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, "0")}-01`;
    if (!buckets[monthKey]) buckets[monthKey] = [];
    buckets[monthKey].push(row);
  });
  return Object.entries(buckets).reduce((acc, [monthKey, monthItems]) => {
    acc[monthKey] = computeMetricValue(monthItems, metricKey);
    return acc;
  }, {});
}

function computeVarianceScore(values) {
  if (!values.length) return 1;
  const mean = values.reduce((acc, value) => acc + value, 0) / values.length;
  const meanAbs = Math.abs(mean) || 1;
  const variance = values.reduce((acc, value) => acc + ((value - mean) ** 2), 0) / values.length;
  return Math.sqrt(variance) / meanAbs;
}

function computeConfidenceEvidence(topItems, bottomItems, metricKey, topValue, bottomValue, delta) {
  const comparedItems = [...topItems, ...bottomItems];
  const metricFields = getMetricFieldKeys(metricKey);
  const totalCompared = comparedItems.length || 1;
  const missingCount = comparedItems.filter((row) => metricFields.some((fieldKey) => !Number.isFinite(row[fieldKey]))).length;
  const missingRate = missingCount / totalCompared;
  const sampleSize = Math.min(topItems.length, bottomItems.length);

  const topMonthly = buildMonthlyMetricMap(topItems, metricKey);
  const bottomMonthly = buildMonthlyMetricMap(bottomItems, metricKey);
  const overlapMonths = Object.keys(topMonthly).filter((monthKey) => Object.prototype.hasOwnProperty.call(bottomMonthly, monthKey));
  const periods = overlapMonths.length;
  const monthlyDeltas = overlapMonths.map((monthKey) => topMonthly[monthKey] - bottomMonthly[monthKey]);
  const varianceScore = computeVarianceScore(monthlyDeltas);
  const deltaPercent = Math.abs(delta) / Math.max(Math.abs(topValue) + Math.abs(bottomValue), 1);

  const reasonCandidates = [];
  if (sampleSize < 5) {
    reasonCandidates.push({ weight: 40, text: "Limited sample size in this segment." });
  }
  if (deltaPercent < 0.08) {
    reasonCandidates.push({ weight: 25, text: "Difference is small relative to overall volume." });
  }
  if (periods < 3) {
    reasonCandidates.push({ weight: 20, text: "Not enough time periods to confirm a stable trend." });
  }
  if (missingRate > 0.1) {
    reasonCandidates.push({ weight: 35, text: "Missing data reduces reliability of the comparison." });
  }
  if (varianceScore > 0.75) {
    reasonCandidates.push({ weight: 15, text: "High volatility makes the trend less consistent." });
  }

  let score = 100;
  reasonCandidates.forEach((reason) => {
    score -= reason.weight;
  });

  let confidenceLevel = "high";
  if (score < 70) confidenceLevel = "medium";
  if (score < 40) confidenceLevel = "low";

  return {
    confidenceLevel,
    confidenceReasons: reasonCandidates
      .sort((a, b) => b.weight - a.weight)
      .slice(0, 2)
      .map((reason) => reason.text),
    confidenceDetails: {
      sampleSize,
      deltaPercent,
      periods,
      missingRate,
      varianceScore,
    },
  };
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
  const { start, end } = computePeriods(rows, 30);
  const currentRows = rows.filter((row) => row.date >= start && row.date <= end);
  const buckets = {};
  currentRows.forEach((row) => {
    const key = row[dimension];
    if (!buckets[key]) buckets[key] = [];
    buckets[key].push(row);
  });

  const breakdown = Object.entries(buckets).map(([key, items]) => ({
    key,
    items,
    value: computeMetricValue(items, metricKey),
    sampleSize: items.length,
  })).sort((a, b) => b.value - a.value);

  if (!breakdown.length) return [];
  const top = breakdown[0];
  const bottom = breakdown[breakdown.length - 1];
  const metricConfig = METRICS.find((m) => m.key === metricKey);
  const metricLabel = metricConfig?.label || metricKey;
  const metricUnit = metricConfig?.unit;
  const delta = top.value - bottom.value;
  const deltaLabel = metricKey === "retention"
    ? `${(delta * 100).toFixed(1)} pp`
    : formatValue(delta, metricUnit);
  const confidenceEvidence = computeConfidenceEvidence(top.items, bottom.items, metricKey, top.value, bottom.value, delta);

  return [
    {
      id: `${metricKey}:${dimension}:${top.key}:${bottom.key}`,
      claim: `${metricLabel} is higher in ${top.key} than ${bottom.key}.`,
      evidence: `Delta ${deltaLabel} (${top.key} vs ${bottom.key}) · n=${top.sampleSize}/${bottom.sampleSize}`,
      driver: `Top segment: ${top.key} (${formatValue(top.value, metricUnit)})`,
      action: `Review ${dimension} drivers to replicate ${top.key} performance.`,
      confidence: confidenceEvidence.confidenceLevel,
      confidenceLevel: confidenceEvidence.confidenceLevel,
      confidenceReasons: confidenceEvidence.confidenceReasons,
      confidenceDetails: confidenceEvidence.confidenceDetails,
    },
  ];
}
export { formatValue };

