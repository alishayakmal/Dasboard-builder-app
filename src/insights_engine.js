(function attachPerformanceInsightsEngine(globalScope) {
  function toNumber(value) {
    if (value === null || value === undefined) return null;
    let text = String(value).trim();
    if (!text) return null;
    if (/^\(.*\)$/.test(text)) text = `-${text.slice(1, -1)}`;
    text = text.replace(/[$€£,%\s,_]/g, "");
    if (!text) return null;
    if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(text)) return null;
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function toDate(value) {
    if (value === null || value === undefined) return null;
    const text = String(value).trim();
    if (!text) return null;
    const parsed = new Date(text);
    return Number.isNaN(parsed.valueOf()) ? null : parsed;
  }

  function mean(values) {
    if (!values.length) return 0;
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }

  function variance(values) {
    if (!values.length) return 0;
    const avg = mean(values);
    return values.reduce((sum, value) => sum + ((value - avg) ** 2), 0) / values.length;
  }

  function median(values) {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    if (sorted.length % 2 === 0) return (sorted[mid - 1] + sorted[mid]) / 2;
    return sorted[mid];
  }

  function compact(value) {
    return new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(Number(value || 0));
  }

  function pct(value) {
    return `${(Number(value || 0) * 100).toFixed(1)}%`;
  }

  function profileDataset(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const columns = Array.from(safeRows.reduce((set, row) => {
      Object.keys(row || {}).forEach((key) => set.add(key));
      return set;
    }, new Set()));
    const columnProfiles = {};
    columns.forEach((column) => {
      const values = safeRows.map((row) => row?.[column]);
      const nonEmpty = values.filter((value) => String(value ?? "").trim() !== "");
      const numericValues = nonEmpty.map((value) => toNumber(value)).filter((value) => value !== null);
      const dateValues = nonEmpty.map((value) => toDate(value)).filter(Boolean);
      const numericRate = nonEmpty.length ? numericValues.length / nonEmpty.length : 0;
      const dateRate = nonEmpty.length ? dateValues.length / nonEmpty.length : 0;
      let type = "categorical";
      if (dateRate >= 0.8) type = "date";
      else if (numericRate >= 0.75) type = "numeric";
      columnProfiles[column] = {
        type,
        nullRate: safeRows.length ? (safeRows.length - nonEmpty.length) / safeRows.length : 1,
        uniqueCount: new Set(nonEmpty.map((value) => String(value))).size,
        min: numericValues.length ? Math.min(...numericValues) : null,
        max: numericValues.length ? Math.max(...numericValues) : null,
        variance: numericValues.length ? variance(numericValues) : null,
        numericRate,
        dateRate,
        sampleSize: nonEmpty.length,
      };
    });
    return {
      rowCount: safeRows.length,
      columnCount: columns.length,
      columns,
      columnProfiles,
    };
  }

  function detectMetricRoles(profile) {
    const columns = profile?.columns || [];
    const byName = profile?.columnProfiles || {};
    const numericColumns = columns.filter((column) => byName[column]?.type === "numeric");
    const dateColumns = columns.filter((column) => byName[column]?.type === "date");
    const entityDimensions = columns.filter((column) => {
      const stats = byName[column];
      return stats?.type === "categorical" && stats.uniqueCount >= 2 && stats.uniqueCount <= 50;
    });
    const rateMetrics = numericColumns.filter((column) => {
      const key = String(column || "").toLowerCase();
      const stats = byName[column] || {};
      const bounded = stats.min !== null && stats.max !== null && stats.min >= 0 && stats.max <= 1.2;
      return /rate|ratio|percent|pct|ctr|cvr|conversion|retention/.test(key) || bounded;
    });
    const costInputMetric = numericColumns.find((column) => /spend|cost|expense|budget|input/.test(String(column || "").toLowerCase())) || null;
    const outcomeCandidates = numericColumns
      .filter((column) => column !== costInputMetric)
      .map((column) => {
        const stats = byName[column] || {};
        const score = (Number(stats.variance || 0) * 0.7) + ((1 - Number(stats.nullRate || 0)) * 0.3);
        return { column, score };
      })
      .sort((a, b) => b.score - a.score);
    const primaryOutcomeMetric = outcomeCandidates[0]?.column || numericColumns[0] || null;
    const efficiencyMetric = costInputMetric && primaryOutcomeMetric
      ? { kind: "derived", name: `${primaryOutcomeMetric}_per_${costInputMetric}`, numerator: primaryOutcomeMetric, denominator: costInputMetric }
      : null;
    const timeField = dateColumns.sort((a, b) => (byName[b]?.dateRate || 0) - (byName[a]?.dateRate || 0))[0] || null;
    return {
      timeField,
      primaryOutcomeMetric,
      costInputMetric,
      efficiencyMetric,
      rateMetrics,
      entityDimensions,
      numericColumns,
    };
  }

  function bucketMonth(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  function sumMetric(rows, metric) {
    return rows.reduce((sum, row) => sum + (toNumber(row?.[metric]) || 0), 0);
  }

  function averageMetric(rows, metric) {
    const values = rows.map((row) => toNumber(row?.[metric])).filter((value) => value !== null);
    return values.length ? mean(values) : 0;
  }

  function computeDelta(previous, current) {
    const abs = Number(current || 0) - Number(previous || 0);
    const pctDelta = Math.abs(Number(previous || 0)) > 0 ? abs / Math.abs(previous) : 0;
    return { previous, current, abs, pct: pctDelta };
  }

  function computePeriodDeltas(rows, roles) {
    if (!roles?.timeField) return { available: false, reason: "time_field_missing", periods: [] };
    const grouped = new Map();
    (rows || []).forEach((row) => {
      const dateValue = toDate(row?.[roles.timeField]);
      if (!dateValue) return;
      const bucket = bucketMonth(dateValue);
      const list = grouped.get(bucket) || [];
      list.push(row);
      grouped.set(bucket, list);
    });
    const periods = Array.from(grouped.keys()).sort();
    if (periods.length < 2) return { available: false, reason: "insufficient_periods", periods };
    const currentPeriod = periods[periods.length - 1];
    const previousPeriod = periods[periods.length - 2];
    const currentRows = grouped.get(currentPeriod) || [];
    const previousRows = grouped.get(previousPeriod) || [];
    const output = {
      available: true,
      periods: [previousPeriod, currentPeriod],
      outcome: computeDelta(
        sumMetric(previousRows, roles.primaryOutcomeMetric),
        sumMetric(currentRows, roles.primaryOutcomeMetric)
      ),
    };
    if (roles.costInputMetric) {
      output.cost = computeDelta(
        sumMetric(previousRows, roles.costInputMetric),
        sumMetric(currentRows, roles.costInputMetric)
      );
      const prevEff = (output.outcome.previous && output.cost.previous) ? output.outcome.previous / output.cost.previous : 0;
      const curEff = (output.outcome.current && output.cost.current) ? output.outcome.current / output.cost.current : 0;
      output.efficiency = computeDelta(prevEff, curEff);
    }
    if (roles.rateMetrics?.length) {
      const topRate = roles.rateMetrics[0];
      output.rate = {
        metric: topRate,
        ...computeDelta(averageMetric(previousRows, topRate), averageMetric(currentRows, topRate)),
      };
    }
    return output;
  }

  function aggregateByDimension(rows, dimension, metric, agg = "sum") {
    const map = new Map();
    (rows || []).forEach((row) => {
      const key = String(row?.[dimension] ?? "Unknown").trim() || "Unknown";
      const value = toNumber(row?.[metric]);
      if (value === null) return;
      const current = map.get(key) || { total: 0, count: 0 };
      current.total += value;
      current.count += 1;
      map.set(key, current);
    });
    const output = Array.from(map.entries()).map(([entity, stats]) => ({
      entity,
      value: agg === "avg" ? (stats.count ? stats.total / stats.count : 0) : stats.total,
      total: stats.total,
      count: stats.count,
    }));
    output.sort((a, b) => b.value - a.value);
    return output;
  }

  function detectPatterns(rows, roles, deltas) {
    const patterns = {
      momentum: null,
      contribution: null,
      concentrationRisk: null,
      leakage: null,
      opportunity: null,
      forecast: null,
      fired: [],
      thresholds: {
        concentrationHigh: 0.6,
        concentrationModerate: 0.4,
        opportunityEfficiencyLift: 0.2,
      },
    };
    if (deltas?.available) {
      const outcomePct = deltas?.outcome?.pct || 0;
      const costPct = deltas?.cost?.pct || 0;
      const efficiencyPct = deltas?.efficiency?.pct || 0;
      if (outcomePct > 0.1 && efficiencyPct > 0.05) patterns.momentum = "Acceleration";
      else if (outcomePct > 0.1 && costPct > 0.1) patterns.momentum = "Scale expansion";
      else if (efficiencyPct < -0.05) patterns.momentum = "Efficiency erosion";
      else patterns.momentum = "Stabilization";
      patterns.fired.push("momentum");
    }

    const driverDimension = roles.entityDimensions?.[0] || null;
    if (driverDimension && roles.primaryOutcomeMetric) {
      const entities = aggregateByDimension(rows, driverDimension, roles.primaryOutcomeMetric, "sum");
      if (entities.length) {
        const total = entities.reduce((sum, item) => sum + item.value, 0);
        const top = entities[0];
        const top2 = entities.slice(0, 2);
        const topShare = total > 0 ? top.value / total : 0;
        const top2Share = total > 0 ? top2.reduce((sum, item) => sum + item.value, 0) / total : 0;
        patterns.contribution = { dimension: driverDimension, topEntities: entities.slice(0, 3), total };
        patterns.concentrationRisk = {
          topShare,
          top2Share,
          level: topShare > 0.6 ? "High" : topShare > 0.4 ? "Moderate" : "Low",
        };
        patterns.fired.push("contribution", "concentration");
      }
    }

    if (roles.costInputMetric && roles.entityDimensions?.length) {
      const dim = roles.entityDimensions[0];
      const costAgg = aggregateByDimension(rows, dim, roles.costInputMetric, "sum");
      const outcomeAgg = aggregateByDimension(rows, dim, roles.primaryOutcomeMetric, "sum");
      const byEntity = new Map();
      costAgg.forEach((item) => byEntity.set(item.entity, { cost: item.value, outcome: 0 }));
      outcomeAgg.forEach((item) => {
        const current = byEntity.get(item.entity) || { cost: 0, outcome: 0 };
        current.outcome = item.value;
        byEntity.set(item.entity, current);
      });
      const combined = Array.from(byEntity.entries()).map(([entity, values]) => ({
        entity,
        cost: values.cost || 0,
        outcome: values.outcome || 0,
        efficiency: values.cost > 0 ? values.outcome / values.cost : 0,
      }));
      const totalCost = combined.reduce((sum, item) => sum + item.cost, 0);
      const totalOutcome = combined.reduce((sum, item) => sum + item.outcome, 0);
      const leakageCandidate = combined
        .map((item) => ({
          ...item,
          costShare: totalCost > 0 ? item.cost / totalCost : 0,
          outcomeShare: totalOutcome > 0 ? item.outcome / totalOutcome : 0,
        }))
        .filter((item) => item.costShare > item.outcomeShare + 0.1)
        .sort((a, b) => (b.costShare - b.outcomeShare) - (a.costShare - a.outcomeShare))[0] || null;
      if (leakageCandidate) {
        patterns.leakage = leakageCandidate;
        patterns.fired.push("leakage");
      }

      const efficiencies = combined.map((item) => item.efficiency).filter((value) => value > 0);
      const med = efficiencies.length ? median(efficiencies) : 0;
      const opportunities = combined.filter((item) => item.efficiency > med * 1.2).sort((a, b) => b.efficiency - a.efficiency).slice(0, 3);
      if (opportunities.length) {
        patterns.opportunity = { medianEfficiency: med, entities: opportunities, dimension: dim };
        patterns.fired.push("opportunity");
      }

      const baselineCost = totalCost;
      const baselineOutcome = totalOutcome;
      const baselineEfficiency = baselineCost > 0 ? baselineOutcome / baselineCost : 0;
      if (baselineCost > 0 && baselineEfficiency > 0) {
        patterns.forecast = [0.1, 0.2, 0.3].map((increase) => ({
          inputIncrease: increase,
          projectedInput: baselineCost * (1 + increase),
          projectedOutcome: baselineOutcome + (baselineCost * increase * baselineEfficiency),
        }));
        patterns.fired.push("forecast");
      }
    }
    return patterns;
  }

  function detectDomain(profile, roles) {
    const columns = profile?.columns || [];
    const text = columns.join(" ").toLowerCase();
    const domainKeywords = {
      marketing: ["campaign", "impression", "click", "ctr", "cpc", "creative", "ad", "roas", "cpa"],
      finance: ["revenue", "margin", "expense", "cost", "budget", "invoice", "profit"],
      product: ["session", "retention", "active", "feature", "adoption", "engagement"],
      operations: ["shipment", "inventory", "fulfillment", "order", "delivery", "capacity"],
    };
    let bestDomain = "general";
    let bestScore = 0;
    let total = 0;
    Object.entries(domainKeywords).forEach(([domain, keywords]) => {
      const score = keywords.reduce((sum, keyword) => sum + (text.includes(keyword) ? 1 : 0), 0);
      total += score;
      if (score > bestScore) {
        bestScore = score;
        bestDomain = domain;
      }
    });
    if (bestScore === 0) return { domain: "general", confidence: 0.5 };
    const confidence = Math.min(1, 0.5 + (bestScore / Math.max(total || bestScore, 1)) * 0.5);
    if (bestDomain === "marketing" && roles?.costInputMetric && roles?.rateMetrics?.length) {
      return { domain: bestDomain, confidence: Math.max(confidence, 0.82) };
    }
    return { domain: bestDomain, confidence };
  }

  function composeNarrative(roles, patterns, deltas, context = {}) {
    const outcome = roles?.primaryOutcomeMetric || "outcome";
    const input = roles?.costInputMetric || "input";
    const hasTime = Boolean(deltas?.available);
    const hasCost = Boolean(roles?.costInputMetric);
    const sections = [];

    const summary = patterns?.contribution?.topEntities?.length
      ? `${patterns.contribution.topEntities[0].entity} leads ${outcome} with ${compact(patterns.contribution.topEntities[0].value)}.`
      : `Current snapshot highlights ${outcome} distribution across available entities.`;
    sections.push({ title: "Executive Summary", bullets: [summary] });

    if (hasTime && patterns?.momentum) {
      sections.push({
        title: "Momentum",
        bullets: [`${patterns.momentum}. Outcome moved ${pct(deltas?.outcome?.pct)} period over period.`],
      });
    } else {
      sections.push({ title: "Momentum", bullets: ["Time trend unavailable. Snapshot analysis only."] });
    }

    const keyBullets = [`Outcome (${outcome}): ${compact(deltas?.outcome?.current || patterns?.contribution?.total || 0)}.`];
    if (hasCost && deltas?.cost) keyBullets.push(`Input (${input}): ${compact(deltas.cost.current)}.`);
    if (hasCost && deltas?.efficiency) keyBullets.push(`Efficiency: ${deltas.efficiency.current.toFixed(3)} ${outcome}/${input}.`);
    if (deltas?.rate?.metric) keyBullets.push(`${deltas.rate.metric}: ${pct(deltas.rate.current)}.`);
    sections.push({ title: "Key Metrics", bullets: keyBullets });

    const winners = patterns?.contribution?.topEntities?.slice(0, 3) || [];
    sections.push({
      title: "Winners",
      bullets: winners.length
        ? winners.map((item) => `${item.entity}: ${compact(item.value)} ${outcome}.`)
        : ["No stable entity ranking available."],
    });

    const spotlights = [];
    if (patterns?.concentrationRisk) {
      spotlights.push(`Concentration risk ${patterns.concentrationRisk.level}. Top entity share ${pct(patterns.concentrationRisk.topShare)}.`);
    }
    if (patterns?.leakage) {
      spotlights.push(`${patterns.leakage.entity} shows high input share ${pct(patterns.leakage.costShare)} with lower outcome share ${pct(patterns.leakage.outcomeShare)}.`);
    }
    sections.push({ title: "Spotlight", bullets: spotlights.length ? spotlights : ["No critical imbalance detected."] });

    sections.push({
      title: "Scale Opportunities",
      bullets: patterns?.opportunity?.entities?.length
        ? patterns.opportunity.entities.map((item) => `${item.entity} efficiency ${item.efficiency.toFixed(3)} exceeds median by 20%+.`)
        : ["No qualified efficiency outliers detected."],
    });

    sections.push({
      title: "Watch List",
      bullets: patterns?.leakage
        ? [`Watch ${patterns.leakage.entity} for input-output imbalance.`]
        : ["Monitor entities with rapid input growth and flat outcome."],
    });

    sections.push({
      title: "Forecast Outlook",
      bullets: hasCost && patterns?.forecast
        ? patterns.forecast.map((item) => `+${Math.round(item.inputIncrease * 100)}% input -> projected outcome ${compact(item.projectedOutcome)}.`)
        : ["Forecast unavailable without a reliable input metric."],
    });

    const nextSteps = [];
    if (patterns?.concentrationRisk?.level === "High") nextSteps.push("Diversify outcome growth beyond the top entity.");
    if (patterns?.opportunity?.entities?.length) nextSteps.push("Scale high-efficiency entities incrementally and measure return.");
    if (patterns?.leakage) nextSteps.push("Audit low-efficiency entities for leakage drivers.");
    sections.push({ title: "Next Steps", bullets: nextSteps.length ? nextSteps.slice(0, 3) : ["Continue monitoring baseline performance and validate assumptions."] });

    const words = sections.flatMap((section) => [section.title, ...(section.bullets || [])]).join(" ").split(/\s+/).filter(Boolean);
    if (words.length > 300) {
      sections.forEach((section) => {
        section.bullets = (section.bullets || []).slice(0, 2);
      });
    }
    return { sections, wordCount: words.length, context };
  }

  function generatePerformanceIntelligence(rows) {
    const profile = profileDataset(rows || []);
    const roles = detectMetricRoles(profile);
    const domain = detectDomain(profile, roles);
    const deltas = computePeriodDeltas(rows || [], roles);
    const patterns = detectPatterns(rows || [], roles, deltas);
    const narrative = composeNarrative(roles, patterns, deltas, { domain });
    const narrativeText = (narrative.sections || [])
      .flatMap((section) => [section.title, ...(section.bullets || [])])
      .join(" ");
    return {
      profile,
      roles,
      domain,
      deltas,
      patterns,
      narrative,
      narrativeText,
      diagnostics: {
        sampleSize: profile.rowCount,
        rolesDetected: roles,
        rulesTriggered: patterns.fired || [],
        thresholds: patterns.thresholds || {},
      },
    };
  }

  const api = {
    profileDataset,
    detectMetricRoles,
    computePeriodDeltas,
    detectPatterns,
    composeNarrative,
    detectDomain,
    generatePerformanceIntelligence,
  };

  globalScope.PerformanceInsightsEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
