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

  const buildKpi = ({ key, label, value, priorValue, deltaType, definition }) => {
    const rawDelta = value - priorValue;
    const delta = deltaType === "percentagePoints"
      ? rawDelta
      : (priorValue ? rawDelta / priorValue : 0);
    const direction = rawDelta > 0 ? "up" : rawDelta < 0 ? "down" : "flat";
    return {
      key,
      label,
      value,
      priorValue,
      deltaRaw: rawDelta,
      delta,
      deltaType,
      direction,
      window: "Last 30 days vs prior 30 days",
      definition,
    };
  };

  return [
    buildKpi({
      key: "revenue",
      label: "Revenue",
      value: revenueCurrent,
      priorValue: revenuePrior,
      deltaType: "percent",
      definition: METRICS.find((m) => m.key === "revenue").definition,
    }),
    buildKpi({
      key: "retention",
      label: "Retention",
      value: retentionCurrent,
      priorValue: retentionPrior,
      deltaType: "percentagePoints",
      definition: METRICS.find((m) => m.key === "retention").definition,
    }),
    buildKpi({
      key: "activeUsers",
      label: "Active users",
      value: activeCurrent,
      priorValue: activePrior,
      deltaType: "percent",
      definition: METRICS.find((m) => m.key === "activeUsers").definition,
    }),
    buildKpi({
      key: "pipeline",
      label: "Pipeline",
      value: pipelineCurrent,
      priorValue: pipelinePrior,
      deltaType: "percent",
      definition: METRICS.find((m) => m.key === "pipeline").definition,
    }),
  ];
}

function toDayKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function movingAverage(values, windowSize = 3) {
  if (!Array.isArray(values) || values.length === 0) return [];
  return values.map((_, index) => {
    const start = Math.max(0, index - (windowSize - 1));
    const slice = values.slice(start, index + 1);
    const sum = slice.reduce((acc, value) => acc + value, 0);
    return sum / slice.length;
  });
}

export function computeKpiMiniTrend(rows, metricKey, windowDays = 30, options = {}) {
  const { smoothSparkline = true } = options;
  const hasDateField = rows.some((row) => row.date instanceof Date && !Number.isNaN(row.date.getTime()));
  if (!hasDateField) {
    return {
      hasDateField: false,
      currentSeries: [],
      priorSeries: [],
      currentLabels: [],
      priorLabels: [],
    };
  }
  const { start, end, priorStart, priorEnd } = computePeriods(rows, windowDays);
  const bucketRange = (rangeStart, rangeEnd) => {
    const days = [];
    const cursor = new Date(rangeStart);
    while (cursor <= rangeEnd) {
      days.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    const dayKeys = days.map((day) => toDayKey(day));
    const buckets = Object.fromEntries(dayKeys.map((key) => [key, []]));
    rows.forEach((row) => {
      if (row.date < rangeStart || row.date > rangeEnd) return;
      const key = toDayKey(row.date);
      if (!buckets[key]) buckets[key] = [];
      buckets[key].push(row);
    });
    const series = dayKeys.map((key) => computeMetricValue(buckets[key] || [], metricKey));
    return {
      labels: dayKeys,
      series: smoothSparkline ? movingAverage(series, 3) : series,
    };
  };
  const current = bucketRange(start, end);
  const prior = bucketRange(priorStart, priorEnd);

  return {
    hasDateField: true,
    currentSeries: current.series,
    priorSeries: prior.series,
    currentLabels: current.labels,
    priorLabels: prior.labels,
  };
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
  const sortedMonths = buildRecentMonthWindow(Array.from(allMonthKeys).sort(), 12);

  return Object.entries(groups).map(([group, months]) => {
    const points = sortedMonths.map((month) => {
      const items = months[month] || [];
      const value = items.length ? computeMetricValue(items, metricKey) : 0;
      return { month, monthTs: new Date(month).getTime(), value };
    });
    return { group, points };
  });
}

function buildRecentMonthWindow(sortedMonthKeys, desiredMonths = 12) {
  if (!sortedMonthKeys.length) return [];
  const latest = new Date(sortedMonthKeys[sortedMonthKeys.length - 1]);
  const end = new Date(latest.getFullYear(), latest.getMonth(), 1);
  const start = new Date(end.getFullYear(), end.getMonth() - (desiredMonths - 1), 1);
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

function computeConfidence({ periods, sampleSize, deltaPercent, variance, missingRate = 0 }) {
  const reasons = [];
  let level = "high";

  if (periods < 3) {
    level = "low";
    reasons.push("Not enough time periods to confirm a stable trend.");
  }

  if (sampleSize < 5) {
    level = "low";
    reasons.push("Limited sample size in this segment.");
  }

  if (missingRate > 0.1) {
    if (level === "high") level = "medium";
    reasons.push("Missing data reduces reliability of the comparison.");
  }

  if (variance > 0.4 && level !== "low") {
    level = "medium";
    reasons.push("Trend shows volatility across periods.");
  }

  if (deltaPercent < 0.1 && level !== "low") {
    level = "medium";
    reasons.push("Performance gap is relatively small.");
  }

  return {
    level,
    reasons: reasons.slice(0, 3),
    metrics: {
      periods,
      sampleSize,
      deltaPercent,
      variance,
      missingRate,
    },
  };
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

  const confidence = computeConfidence({
    periods,
    sampleSize,
    deltaPercent,
    variance: varianceScore,
    missingRate,
  });

  return {
    confidence,
    confidenceLevel: confidence.level,
    confidenceReasons: confidence.reasons.slice(0, 2),
    confidenceDetails: {
      ...confidence.metrics,
      varianceScore,
    },
  };
}


function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function buildAction({ metricKey, topSegment, driverCombination, confidence, deltaStrength }) {
  const metricKeyLower = String(metricKey || "").toLowerCase();
  const comboText = driverCombination || topSegment || "the top driver";

  if (confidence?.level === "low") {
    return `Validate whether ${comboText} remains the primary driver with additional time periods before reallocating budget.`;
  }

  if (metricKeyLower === "retention") {
    return `Identify what is unique about ${comboText} and apply the same lifecycle tactics to lift retention outside ${topSegment}.`;
  }

  if (metricKeyLower === "revenue") {
    return `Rebalance investment toward ${comboText} while protecting efficiency, then test expansion into the next best segment.`;
  }

  if (metricKeyLower === "pipeline") {
    return `Inspect funnel stage conversion for ${comboText} and replicate the strongest stage improvements in lower performing segments.`;
  }

  if ((deltaStrength || 0) < 0.1) {
    return `Monitor ${comboText} and run a controlled test before scaling changes due to a relatively small performance gap.`;
  }

  return `Investigate drivers behind ${comboText} and run a controlled test to validate impact.`;
}

function formatComboNarrative(dimensions, valuesByDimension) {
  const has = (key) => Object.prototype.hasOwnProperty.call(valuesByDimension, key);
  if (has("plan") && has("region") && has("channel")) {
    return `${valuesByDimension.plan} in ${valuesByDimension.region} via ${valuesByDimension.channel}`;
  }
  if (has("plan") && has("region")) {
    return `${valuesByDimension.plan} in ${valuesByDimension.region}`;
  }
  if (has("plan") && has("channel")) {
    return `${valuesByDimension.plan} via ${valuesByDimension.channel}`;
  }
  return dimensions.map((dim) => valuesByDimension[dim]).join(" + ");
}

function buildDimensionCombos(primaryDimension, candidateDimensions) {
  const combos = [];
  candidateDimensions.forEach((dim) => {
    combos.push([primaryDimension, dim]);
  });
  for (let i = 0; i < candidateDimensions.length; i += 1) {
    for (let j = i + 1; j < candidateDimensions.length; j += 1) {
      combos.push([primaryDimension, candidateDimensions[i], candidateDimensions[j]]);
    }
  }
  return combos;
}

function aggregateComboRows(rows, metricKey, dimensions, winnerDimension, winnerValue) {
  const buckets = {};
  rows.forEach((row) => {
    if (row[winnerDimension] !== winnerValue) return;
    const valuesByDimension = {};
    dimensions.forEach((dim) => {
      valuesByDimension[dim] = row[dim] ?? "Unknown";
    });
    const key = dimensions.map((dim) => valuesByDimension[dim]).join("||");
    if (!buckets[key]) {
      buckets[key] = {
        key,
        dimensions,
        valuesByDimension,
        items: [],
      };
    }
    buckets[key].items.push(row);
  });
  return Object.values(buckets).map((bucket) => ({
    ...bucket,
    value: computeMetricValue(bucket.items, metricKey),
  }));
}

function computeTopDriverCombo({
  rows,
  currentRows,
  priorRows,
  metricKey,
  primaryDimension,
  winnerKey,
  winnerValue,
  confidenceLevel,
  metricUnit,
}) {
  const candidatePool = ["region", "channel", "device", "creative"]
    .filter((dim) => dim !== primaryDimension && rows.some((row) => Object.prototype.hasOwnProperty.call(row, dim)));
  const dimensionSets = buildDimensionCombos(primaryDimension, candidatePool);
  if (!dimensionSets.length) {
    return {
      narrative: "No dominant driver identified.",
      breakdownTable: [],
      driverNarrative: "No dominant driver identified.",
      dominant: false,
      liftShare: 0,
      filtersUsed: `${primaryDimension}=${winnerKey}`,
    };
  }

  let bestCandidate = null;

  dimensionSets.forEach((dimensions) => {
    const currentBuckets = aggregateComboRows(currentRows, metricKey, dimensions, primaryDimension, winnerKey);
    const priorBuckets = aggregateComboRows(priorRows, metricKey, dimensions, primaryDimension, winnerKey);
    const priorMap = Object.fromEntries(priorBuckets.map((bucket) => [bucket.key, bucket]));
    const currentMap = Object.fromEntries(currentBuckets.map((bucket) => [bucket.key, bucket]));
    const allKeys = Array.from(new Set([...Object.keys(currentMap), ...Object.keys(priorMap)]));
    if (allKeys.length > 50) return;

    const merged = allKeys.map((key) => {
      const currentBucket = currentMap[key];
      const priorBucket = priorMap[key];
      const currentValue = currentBucket?.value || 0;
      const priorValue = priorBucket?.value || 0;
      const comboDelta = currentValue - priorValue;
      const valuesByDimension = currentBucket?.valuesByDimension || priorBucket?.valuesByDimension || {};
      return {
        key,
        dimensions,
        valuesByDimension,
        label: formatComboNarrative(dimensions, valuesByDimension),
        currentValue,
        priorValue,
        comboDelta,
      };
    }).sort((a, b) => b.currentValue - a.currentValue).slice(0, 10);

    const totalDeltaAcrossAllCombos = merged.reduce((acc, item) => acc + Math.max(0, item.comboDelta), 0);
    if (Math.abs(totalDeltaAcrossAllCombos) < 1e-9) return;

    merged.forEach((item) => {
      const shareOfWinner = winnerValue ? item.currentValue / winnerValue : 0;
      const liftShare = item.comboDelta > 0 ? item.comboDelta / totalDeltaAcrossAllCombos : 0;
      const score = (0.6 * liftShare) + (0.4 * shareOfWinner);
      const candidate = {
        ...item,
        shareOfWinner,
        liftShare,
        score,
        winnerKey,
        winnerValue,
        metricUnit,
        filtersUsed: `${primaryDimension}=${winnerKey}; dims=${dimensions.join(",")}`,
      };
      if (!bestCandidate || candidate.score > bestCandidate.score) {
        bestCandidate = candidate;
      }
    });
  });

  if (!bestCandidate || bestCandidate.liftShare <= 0.05) {
    return {
      narrative: "No dominant driver identified.",
      breakdownTable: [],
      driverNarrative: "No dominant driver identified.",
      dominant: false,
      liftShare: 0,
      filtersUsed: `${primaryDimension}=${winnerKey}`,
    };
  }

  const chosenDims = bestCandidate.dimensions;
  const currentBuckets = aggregateComboRows(currentRows, metricKey, chosenDims, primaryDimension, winnerKey);
  const priorBuckets = aggregateComboRows(priorRows, metricKey, chosenDims, primaryDimension, winnerKey);
  const priorMap = Object.fromEntries(priorBuckets.map((bucket) => [bucket.key, bucket]));
  const chosenMergedAll = Array.from(new Set([
    ...currentBuckets.map((bucket) => bucket.key),
    ...priorBuckets.map((bucket) => bucket.key),
  ])).map((key) => {
    const currentBucket = currentBuckets.find((bucket) => bucket.key === key);
    const priorBucket = priorBuckets.find((bucket) => bucket.key === key);
    return {
      key,
      currentValue: currentBucket?.value || 0,
      priorValue: priorBucket?.value || 0,
    };
  });
  const chosenTotalDelta = chosenMergedAll.reduce((acc, item) => acc + Math.max(0, item.currentValue - item.priorValue), 0) || 1;
  const table = currentBuckets
    .map((bucket) => {
      const priorValue = priorMap[bucket.key]?.value || 0;
      const delta = bucket.value - priorValue;
      const rowLiftShare = delta > 0 ? delta / chosenTotalDelta : 0;
      return {
        combination: formatComboNarrative(chosenDims, bucket.valuesByDimension),
        current: bucket.value,
        currentLabel: formatValue(bucket.value, metricUnit),
        prior: priorValue,
        priorLabel: formatValue(priorValue, metricUnit),
        delta,
        deltaLabel: metricUnit === "rate" ? `${(delta * 100).toFixed(1)} pp` : formatValue(delta, metricUnit),
        liftShare: rowLiftShare,
        liftShareLabel: `${Math.round(rowLiftShare * 100)}%`,
      };
    })
    .sort((a, b) => b.current - a.current)
    .slice(0, 5);

  const leadLabel = formatComboNarrative(bestCandidate.dimensions, bestCandidate.valuesByDimension);
  const prefix = confidenceLevel === "low" ? "Driver candidate, validate with more periods: " : "";
  return {
    narrative: `${prefix}${leadLabel} contributes ${Math.round(bestCandidate.liftShare * 100)}% of the lift.`,
    driverNarrative: leadLabel,
    dominant: true,
    liftShare: bestCandidate.liftShare,
    breakdownTable: table,
    filtersUsed: bestCandidate.filtersUsed,
  };
}

function buildEvidenceDetail({ breakdown, top, bottom, metricKey, metricUnit, confidenceEvidence, driverInfo }) {
  const totalValue = breakdown.reduce((acc, item) => acc + Math.max(0, item.value), 0) || 1;
  const comparisonTable = breakdown.slice(0, 5).map((item, index) => ({
    segment: item.key,
    value: item.value,
    valueLabel: formatValue(item.value, metricUnit),
    share: item.value / totalValue,
    shareLabel: `${Math.round((item.value / totalValue) * 100)}%`,
    rank: index + 1,
  }));

  const topMonthly = buildMonthlyMetricMap(top.items, metricKey);
  const bottomMonthly = buildMonthlyMetricMap(bottom.items, metricKey);
  const overlapMonths = Object.keys(topMonthly)
    .filter((monthKey) => Object.prototype.hasOwnProperty.call(bottomMonthly, monthKey))
    .sort();
  const leadPeriods = overlapMonths.filter((monthKey) => topMonthly[monthKey] > bottomMonthly[monthKey]).length;
  const comparedPeriods = overlapMonths.length;
  const relativeLift = Math.abs(bottom.value) > 0 ? ((top.value - bottom.value) / Math.abs(bottom.value)) : 0;
  const consistencyScore = clamp(
    Math.round((1 / (1 + (confidenceEvidence.confidenceDetails.varianceScore || 0))) * 100),
    0,
    100
  );

  const metricLabel = metricKey === "activeUsers" ? "active users" : metricKey;
  const contributionSummary = `${top.key} contributes ${Math.round(relativeLift * 100)}% more ${metricLabel} than ${bottom.key} and leads in ${leadPeriods} of ${Math.max(comparedPeriods, 1)} overlapping periods.`;

  return {
    comparisonTable,
    contributionSummary,
    driverNarrative: driverInfo?.driverNarrative || "No dominant driver identified.",
    driverDominant: Boolean(driverInfo?.dominant),
    driverBreakdownTable: driverInfo?.breakdownTable || [],
    driverFiltersUsed: driverInfo?.filtersUsed || "",
    trendSummary: {
      periodsAnalysed: confidenceEvidence.confidenceDetails.periods,
      variance: confidenceEvidence.confidenceDetails.varianceScore || 0,
      leadPeriods,
      comparedPeriods,
    },
    consistencyScore,
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
  const { priorStart, priorEnd } = computePeriods(rows, 30);
  const priorRows = rows.filter((row) => row.date >= priorStart && row.date <= priorEnd);
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
  const driverInfo = computeTopDriverCombo({
    rows,
    currentRows,
    priorRows,
    metricKey,
    primaryDimension: dimension,
    winnerKey: top.key,
    winnerValue: top.value,
    confidenceLevel: confidenceEvidence.confidence.level,
    metricUnit,
  });
  const evidence = buildEvidenceDetail({
    breakdown,
    top,
    bottom,
    metricKey,
    metricUnit,
    confidenceEvidence,
    driverInfo,
  });
  const actionSummary = buildAction({
    metricKey,
    topSegment: top.key,
    driverCombination: driverInfo.driverNarrative,
    confidence: confidenceEvidence.confidence,
    deltaStrength: confidenceEvidence.confidenceDetails.deltaPercent,
  });

  return [
    {
      id: `${metricKey}:${dimension}:${top.key}:${bottom.key}`,
      headline: `${metricLabel} is higher in ${top.key} than ${bottom.key}.`,
      claim: `${metricLabel} is higher in ${top.key} than ${bottom.key}.`,
      delta,
      metricKey,
      dimensionKey: dimension,
      deltaSummary: `Delta ${deltaLabel} (${top.key} vs ${bottom.key}) · n=${top.sampleSize}/${bottom.sampleSize}`,
      topSegment: top.key,
      driverCombination: driverInfo.driverNarrative,
      driverNarrative: driverInfo.narrative,
      driver: driverInfo.narrative,
      action: actionSummary,
      confidence: confidenceEvidence.confidence,
      confidenceLevel: confidenceEvidence.confidence.level,
      confidenceReasons: confidenceEvidence.confidence.reasons.slice(0, 2),
      confidenceDetails: confidenceEvidence.confidenceDetails,
      evidence,
    },
  ];
}
export { formatValue };

