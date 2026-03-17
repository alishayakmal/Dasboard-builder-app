(function attachPerformanceInsightsEngine(globalScope) {
  function toNumber(value) {
    if (value === null || value === undefined) return null;
    let text = String(value).trim();
    if (!text) return null;
    if (/^\(.*\)$/.test(text)) text = `-${text.slice(1, -1)}`;
    const hasPercent = /%/.test(text);
    text = text.replace(/[$€£,%\s,_]/g, "");
    if (!text) return null;
    if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(text)) return null;
    const parsed = Number(text);
    if (!Number.isFinite(parsed)) return null;
    return hasPercent ? parsed / 100 : parsed;
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
    return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
  }

  function compact(value) {
    return new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(Number(value || 0));
  }

  function pct(value) {
    return `${(Number(value || 0) * 100).toFixed(1)}%`;
  }

  function isPlaceholderValue(value) {
    const key = String(value ?? "").trim().toLowerCase();
    if (!key) return true;
    return [
      "unknown",
      "unassigned",
      "n/a",
      "na",
      "null",
      "none",
      "(not set)",
      "not set",
      "undefined",
      "other",
      "others",
    ].includes(key);
  }

  function isLowSignalKey(key) {
    return String(key || "").split("|").some((part) => isPlaceholderValue(part));
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
        rowCount: safeRows.length,
        nullRate: safeRows.length ? (safeRows.length - nonEmpty.length) / safeRows.length : 1,
        uniqueCount: new Set(nonEmpty.map((value) => String(value))).size,
        uniqueRate: nonEmpty.length ? new Set(nonEmpty.map((value) => String(value))).size / nonEmpty.length : 0,
        min: numericValues.length ? Math.min(...numericValues) : null,
        max: numericValues.length ? Math.max(...numericValues) : null,
        variance: numericValues.length ? variance(numericValues) : null,
        integerOnly: numericValues.length ? numericValues.every((value) => Number.isInteger(value)) : false,
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

  function classifyNumericMetric(name, stats) {
    const key = String(name || "").toLowerCase();
    const min = Number(stats?.min ?? 0);
    const max = Number(stats?.max ?? 0);
    const boundedRate = Number.isFinite(min) && Number.isFinite(max) && min >= 0 && max <= 1.2;
    const isRateByName = /(?:^|[\s_-])(ctr|cvr|rate|ratio|percent|pct)(?:$|[\s_-])/.test(key)
      || /\bconversion\s*rate\b/.test(key)
      || /\bclick[\s-]*through[\s-]*rate\b/.test(key);
    const isRate = isRateByName || boundedRate;
    const isEfficiency = /(?:^|[\s_-])(cpa|cpc|cpm|ecpa|ecpc|ecpm|roas|roi)(?:$|[\s_-])|per[\s_-]/.test(key);
    const isInput = /spend|cost|expense|budget|input/.test(key);
    const additive = !isRate && !isEfficiency;
    let role = "volume";
    if (isInput) role = "input";
    else if (isRate) role = "rate";
    else if (isEfficiency) role = "efficiency";
    return { role, additive, isInput, isRate, isEfficiency };
  }

  function isIdentifierLikeField(name, stats, rowCount = 0) {
    const key = String(name || "").trim().toLowerCase();
    const uniqueRate = Number(stats?.uniqueRate || 0);
    const integerOnly = Boolean(stats?.integerOnly);
    const nonBlank = Number(stats?.sampleSize || 0);
    const namePatterns = ["id", "uuid", "guid", "identifier", "key", "code", "reference", "ref", "number", "num"];
    const nameSuggestsIdentifier = namePatterns.some((pattern) => (
      key === pattern
      || key.endsWith(` ${pattern}`)
      || key.endsWith(`_${pattern}`)
      || key.includes(`${pattern} `)
      || key.includes(`_${pattern}`)
    ));
    if (nameSuggestsIdentifier) return true;
    if (rowCount >= 20 && uniqueRate >= 0.9) {
      if (nameSuggestsIdentifier) return true;
      if (integerOnly) {
        const range = (stats?.max !== null && stats?.min !== null) ? Math.abs(Number(stats.max) - Number(stats.min)) : 0;
        if (nonBlank >= 20 && uniqueRate >= 0.98 && range > nonBlank * 10) return true;
      }
    }
    if (integerOnly && stats?.min !== null && stats?.max !== null) {
      const range = Math.abs(Number(stats.max) - Number(stats.min));
      if (nonBlank >= 20 && uniqueRate >= 0.8 && range > nonBlank * 10) return true;
    }
    return false;
  }

  function isContributionDimensionCandidate(name, stats) {
    const key = String(name || "").toLowerCase();
    if (!stats || stats.type !== "categorical") return false;
    if (stats.uniqueCount < 2 || stats.uniqueCount > 80) return false;
    if (isIdentifierLikeField(name, stats, stats?.rowCount || 0)) return false;
    if (stats.numericRate >= 0.35) return false;
    if (stats.nullRate >= 0.75) return false;
    if (/(ctr|cvr|rate|ratio|percent|pct|cpa|cpc|cpm|ecpa|ecpc|ecpm|roas|roi|secondary|primary)/.test(key)) return false;
    if (/(cost|spend|budget|expense|revenue|sales|clicks|impressions|conversions)/.test(key)) return false;
    return true;
  }

  function detectMetricRoles(profile, options = {}) {
    const columns = profile?.columns || [];
    const byName = profile?.columnProfiles || {};
    const numericColumns = columns.filter((column) => byName[column]?.type === "numeric");
    const dateColumns = columns.filter((column) => byName[column]?.type === "date");
    const entityDimensions = columns.filter((column) => isContributionDimensionCandidate(column, byName[column]));
    const metricRoles = {};
    numericColumns.forEach((column) => {
      metricRoles[column] = classifyNumericMetric(column, byName[column]);
    });
    const additiveNumeric = numericColumns
      .filter((column) => metricRoles[column]?.additive)
      .filter((column) => !isIdentifierLikeField(column, byName[column], profile?.rowCount || 0));
    const costInputMetric = additiveNumeric.find((column) => metricRoles[column]?.isInput) || null;

    const outcomeCandidates = additiveNumeric
      .filter((column) => column !== costInputMetric)
      .filter((column) => !metricRoles[column]?.isInput)
      .map((column) => {
        const stats = byName[column] || {};
        const score = (Number(stats.variance || 0) * 0.7) + ((1 - Number(stats.nullRate || 0)) * 0.3);
        return { column, score };
      })
      .sort((a, b) => b.score - a.score);

    const preferredOutcomeMetric = options?.preferredOutcomeMetric && columns.includes(options.preferredOutcomeMetric)
      ? options.preferredOutcomeMetric
      : null;
    let primaryOutcomeMetric = outcomeCandidates[0]?.column || additiveNumeric.find((column) => !metricRoles[column]?.isInput) || null;
    let preferredOutcomeReason = null;
    if (preferredOutcomeMetric) {
      if (metricRoles[preferredOutcomeMetric]?.isInput) {
        preferredOutcomeReason = "selected_metric_is_input";
      } else if (metricRoles[preferredOutcomeMetric]?.additive) {
        primaryOutcomeMetric = preferredOutcomeMetric;
      } else {
        preferredOutcomeReason = "selected_metric_not_additive";
      }
    }
    const outcomeIsNonAdditive = Boolean(
      primaryOutcomeMetric
      && metricRoles[primaryOutcomeMetric]
      && !metricRoles[primaryOutcomeMetric].additive
    );

    const rateMetrics = numericColumns.filter((column) => metricRoles[column]?.isRate);
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
      additiveMetrics: additiveNumeric,
      metricRoles,
      preferredOutcomeMetric,
      preferredOutcomeReason,
      outcomeIsNonAdditive,
    };
  }

  function bucketMonth(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  function aggregateMetric(rows, metric, agg = "sum") {
    const values = (rows || []).map((row) => toNumber(row?.[metric])).filter((value) => value !== null);
    if (!values.length) return 0;
    if (agg === "avg") return mean(values);
    return values.reduce((sum, value) => sum + value, 0);
  }

  function computeDelta(previous, current) {
    const abs = Number(current || 0) - Number(previous || 0);
    const pctDelta = Math.abs(Number(previous || 0)) > 0 ? abs / Math.abs(previous) : 0;
    return { previous, current, abs, pct: pctDelta };
  }

  function computeRateWeighted(rows, numeratorMetric, denominatorMetric) {
    if (!numeratorMetric || !denominatorMetric) return null;
    const numerator = aggregateMetric(rows, numeratorMetric, "sum");
    const denominator = aggregateMetric(rows, denominatorMetric, "sum");
    if (!denominator) return null;
    return numerator / denominator;
  }

  function computePeriodDeltas(rows, roles) {
    if (!roles?.primaryOutcomeMetric) return { available: false, reason: "outcome_metric_missing", periods: [] };
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
        aggregateMetric(previousRows, roles.primaryOutcomeMetric, "sum"),
        aggregateMetric(currentRows, roles.primaryOutcomeMetric, "sum")
      ),
    };

    if (roles.costInputMetric) {
      output.cost = computeDelta(
        aggregateMetric(previousRows, roles.costInputMetric, "sum"),
        aggregateMetric(currentRows, roles.costInputMetric, "sum")
      );
      const prevEff = output.cost.previous ? output.outcome.previous / output.cost.previous : 0;
      const curEff = output.cost.current ? output.outcome.current / output.cost.current : 0;
      output.efficiency = computeDelta(prevEff, curEff);
    }

    const topRate = (roles.rateMetrics || [])[0] || null;
    if (topRate) {
      const key = String(topRate || "").toLowerCase();
      let prevRate = null;
      let curRate = null;
      if (/ctr/.test(key)) {
        prevRate = computeRateWeighted(previousRows, "clicks", "impressions");
        curRate = computeRateWeighted(currentRows, "clicks", "impressions");
      } else if (/cvr/.test(key)) {
        prevRate = computeRateWeighted(previousRows, "conversions", "clicks");
        curRate = computeRateWeighted(currentRows, "conversions", "clicks");
      }
      if (prevRate !== null && curRate !== null) {
        output.rate = { metric: topRate, ...computeDelta(prevRate, curRate) };
      }
    }
    return output;
  }

  function buildSingleDimensionTable(rows, dimension, roles) {
    const byEntity = new Map();
    (rows || []).forEach((row) => {
      const entity = String(row?.[dimension] ?? "Unknown").trim() || "Unknown";
      const current = byEntity.get(entity) || { entity, outcome: 0, input: 0, rows: 0 };
      current.outcome += toNumber(row?.[roles.primaryOutcomeMetric]) || 0;
      current.input += roles.costInputMetric ? (toNumber(row?.[roles.costInputMetric]) || 0) : 0;
      current.rows += 1;
      byEntity.set(entity, current);
    });
    return Array.from(byEntity.values());
  }

  function buildPairTable(rows, dimA, dimB, roles) {
    const byPair = new Map();
    (rows || []).forEach((row) => {
      const a = String(row?.[dimA] ?? "Unknown").trim() || "Unknown";
      const b = String(row?.[dimB] ?? "Unknown").trim() || "Unknown";
      const pairKey = `${a} | ${b}`;
      const current = byPair.get(pairKey) || { pairKey, dimAValue: a, dimBValue: b, outcome: 0, input: 0, rows: 0 };
      current.outcome += toNumber(row?.[roles.primaryOutcomeMetric]) || 0;
      current.input += roles.costInputMetric ? (toNumber(row?.[roles.costInputMetric]) || 0) : 0;
      current.rows += 1;
      byPair.set(pairKey, current);
    });
    return Array.from(byPair.values());
  }

  function topWithOther(items, valueKey, topN) {
    const sorted = [...(items || [])].sort((a, b) => (Number(b?.[valueKey] || 0) - Number(a?.[valueKey] || 0)));
    const head = sorted.slice(0, topN);
    const tail = sorted.slice(topN);
    if (tail.length) {
      const other = { label: "Other" };
      Object.keys(head[0] || {}).forEach((key) => {
        if (typeof head[0][key] === "number") {
          other[key] = tail.reduce((sum, item) => sum + Number(item?.[key] || 0), 0);
        }
      });
      other.entity = "Other";
      other.pairKey = "Other";
      head.push(other);
    }
    return head;
  }

  function evaluatePairEligibility(rows, dimA, dimB) {
    const uniqueA = new Set((rows || []).map((row) => String(row?.[dimA] ?? "Unknown"))).size;
    const uniqueB = new Set((rows || []).map((row) => String(row?.[dimB] ?? "Unknown"))).size;
    const combined = new Set((rows || []).map((row) => `${String(row?.[dimA] ?? "Unknown")} | ${String(row?.[dimB] ?? "Unknown")}`)).size;
    const caps = { dimA: 40, dimB: 40, combined: 180 };
    const eligible = uniqueA <= caps.dimA && uniqueB <= caps.dimB && combined <= caps.combined;
    return {
      eligible,
      uniqueA,
      uniqueB,
      combined,
      caps,
      reason: eligible ? "eligible" : `cap_failed(dimA=${uniqueA}, dimB=${uniqueB}, combined=${combined})`,
    };
  }

  function computeExplanatoryPower(table) {
    const sorted = [...(table || [])].sort((a, b) => Number(b?.outcome || 0) - Number(a?.outcome || 0));
    const total = sorted.reduce((sum, row) => sum + Number(row?.outcome || 0), 0);
    if (!total || !sorted.length) return 0;
    const top = Number(sorted[0]?.outcome || 0);
    const second = Number(sorted[1]?.outcome || 0);
    const incremental = Math.max(0, top - second);
    return incremental / total;
  }

  function chooseBestGrain(rows, roles, options = {}) {
    const dims = roles?.entityDimensions || [];
    const preferredA = options?.preferredDimA;
    const preferredB = options?.preferredDimB;

    const singleCandidates = dims.map((dim) => {
      const table = buildSingleDimensionTable(rows, dim, roles);
      return { kind: "single", dimension: dim, table, explanatoryPower: computeExplanatoryPower(table) };
    }).sort((a, b) => b.explanatoryPower - a.explanatoryPower);

    const bestSingle = singleCandidates[0] || null;

    let pairCandidates = [];
    if (preferredA && preferredB && preferredA !== preferredB && dims.includes(preferredA) && dims.includes(preferredB)) {
      pairCandidates = [[preferredA, preferredB]];
    } else {
      for (let i = 0; i < dims.length; i += 1) {
        for (let j = i + 1; j < dims.length; j += 1) {
          pairCandidates.push([dims[i], dims[j]]);
        }
      }
      if (preferredA && dims.includes(preferredA)) {
        pairCandidates.sort((a, b) => {
          const aHas = a[0] === preferredA || a[1] === preferredA ? 1 : 0;
          const bHas = b[0] === preferredA || b[1] === preferredA ? 1 : 0;
          return bHas - aHas;
        });
      }
    }

    const evaluatedPairs = pairCandidates.map(([dimA, dimB]) => {
      const eligibility = evaluatePairEligibility(rows, dimA, dimB);
      if (!eligibility.eligible) {
        return { kind: "pair", dimA, dimB, eligible: false, eligibility, table: [], explanatoryPower: 0 };
      }
      const table = buildPairTable(rows, dimA, dimB, roles);
      return {
        kind: "pair",
        dimA,
        dimB,
        eligible: true,
        eligibility,
        table,
        explanatoryPower: computeExplanatoryPower(table),
      };
    }).sort((a, b) => b.explanatoryPower - a.explanatoryPower);

    const bestPair = evaluatedPairs.find((item) => item.eligible) || null;
    const bestSinglePower = bestSingle?.explanatoryPower || 0;
    const pairPreferenceThreshold = 0.85;
    const pairAnchoredOnPreferredA = Boolean(bestPair && preferredA && (bestPair.dimA === preferredA || bestPair.dimB === preferredA));
    const shouldUsePair = Boolean(
      bestPair
      && (
        bestSinglePower === 0
        || bestPair.explanatoryPower >= (bestSinglePower * pairPreferenceThreshold)
        || (pairAnchoredOnPreferredA && bestPair.explanatoryPower >= (bestSinglePower * 0.6))
      )
    );
    const chosen = shouldUsePair
      ? {
          grain: "paired",
          table: bestPair.table,
          keyField: "pairKey",
          dimension: `${bestPair.dimA} | ${bestPair.dimB}`,
          dimA: bestPair.dimA,
          dimB: bestPair.dimB,
          explanatoryPower: bestPair.explanatoryPower,
          eligibility: bestPair.eligibility,
          comparedSinglePower: bestSinglePower,
        }
      : {
          grain: "single",
          table: bestSingle?.table || [],
          keyField: "entity",
          dimension: bestSingle?.dimension || null,
          dimA: bestSingle?.dimension || null,
          dimB: null,
          explanatoryPower: bestSingle?.explanatoryPower || 0,
          eligibility: bestPair?.eligibility || null,
          comparedSinglePower: bestSinglePower,
        };

    return {
      chosen,
      bestSingle,
      bestPair,
      evaluatedPairs,
    };
  }

  function evaluateTripleEligibility(rows, dimA, dimB, dimC) {
    const combined = new Set((rows || []).map((row) => [
      String(row?.[dimA] ?? "Unknown"),
      String(row?.[dimB] ?? "Unknown"),
      String(row?.[dimC] ?? "Unknown"),
    ].join(" | "))).size;
    const caps = { combined: 120 };
    const eligible = combined <= caps.combined;
    return {
      eligible,
      combined,
      caps,
      reason: eligible ? "eligible" : `cap_failed(combined=${combined})`,
    };
  }

  function buildTripleTable(rows, dimA, dimB, dimC, roles) {
    const byTriple = new Map();
    (rows || []).forEach((row) => {
      const a = String(row?.[dimA] ?? "Unknown").trim() || "Unknown";
      const b = String(row?.[dimB] ?? "Unknown").trim() || "Unknown";
      const c = String(row?.[dimC] ?? "Unknown").trim() || "Unknown";
      const tripleKey = `${a} | ${b} | ${c}`;
      const current = byTriple.get(tripleKey) || { tripleKey, outcome: 0, input: 0, rows: 0 };
      current.outcome += toNumber(row?.[roles.primaryOutcomeMetric]) || 0;
      current.input += roles.costInputMetric ? (toNumber(row?.[roles.costInputMetric]) || 0) : 0;
      current.rows += 1;
      byTriple.set(tripleKey, current);
    });
    return Array.from(byTriple.values());
  }

  function buildContributionDriverAnalysis(rows, roles, grainChoice, options = {}) {
    const dims = roles?.entityDimensions || [];
    const preferredA = options?.preferredDimA;
    const preferredB = options?.preferredDimB;
    const chosenDims = grainChoice?.chosen?.grain === "paired"
      ? [grainChoice?.chosen?.dimA, grainChoice?.chosen?.dimB].filter(Boolean)
      : [grainChoice?.chosen?.dimA].filter(Boolean);
    const rankedCandidates = [];

    dims.forEach((dim) => {
      const table = buildSingleDimensionTable(rows, dim, roles);
      rankedCandidates.push({
        kind: "single",
        dimensions: [dim],
        table,
        explanatoryPower: computeExplanatoryPower(table),
      });
    });

    for (let i = 0; i < dims.length; i += 1) {
      for (let j = i + 1; j < dims.length; j += 1) {
        const dimA = dims[i];
        const dimB = dims[j];
        const eligibility = evaluatePairEligibility(rows, dimA, dimB);
        if (!eligibility.eligible) continue;
        const table = buildPairTable(rows, dimA, dimB, roles);
        rankedCandidates.push({
          kind: "pair",
          dimensions: [dimA, dimB],
          table,
          explanatoryPower: computeExplanatoryPower(table),
          eligibility,
        });
      }
    }

    for (let i = 0; i < dims.length; i += 1) {
      for (let j = i + 1; j < dims.length; j += 1) {
        for (let k = j + 1; k < dims.length; k += 1) {
          const dimA = dims[i];
          const dimB = dims[j];
          const dimC = dims[k];
          const includesPreferredA = !preferredA || [dimA, dimB, dimC].includes(preferredA);
          const includesPreferredB = !preferredB || [dimA, dimB, dimC].includes(preferredB);
          if (!includesPreferredA && !includesPreferredB) continue;
          const eligibility = evaluateTripleEligibility(rows, dimA, dimB, dimC);
          if (!eligibility.eligible) continue;
          const table = buildTripleTable(rows, dimA, dimB, dimC, roles);
          rankedCandidates.push({
            kind: "triple",
            dimensions: [dimA, dimB, dimC],
            table,
            explanatoryPower: computeExplanatoryPower(table),
            eligibility,
          });
        }
      }
    }

    rankedCandidates.forEach((candidate) => {
      const includesPreferredA = preferredA && candidate.dimensions.includes(preferredA) ? 1 : 0;
      const includesPreferredB = preferredB && candidate.dimensions.includes(preferredB) ? 1 : 0;
      const pairBonus = candidate.kind === "pair" ? 0.09 : 0;
      const tripleBonus = candidate.kind === "triple" ? 0.07 : 0;
      const chosenDimMatches = chosenDims.filter((dim) => candidate.dimensions.includes(dim)).length;
      const chosenDimBonus = chosenDimMatches ? (0.05 * chosenDimMatches) : 0;
      candidate.rankScore = candidate.explanatoryPower + (includesPreferredA * 0.08) + (includesPreferredB * 0.05) + pairBonus + tripleBonus + chosenDimBonus;
    });

    rankedCandidates.sort((a, b) => b.rankScore - a.rankScore);
    const chosen = rankedCandidates[0] || null;
    if (!chosen) {
      return {
        kind: "single",
        dimensions: [],
        driverDimensionLabel: "entity",
        topDrivers: [],
        totalOutcome: 0,
        explanatoryPower: 0,
      };
    }

    const keyField = chosen.kind === "pair"
      ? "pairKey"
      : chosen.kind === "triple"
        ? "tripleKey"
        : "entity";
    const normalized = (chosen.table || []).map((row) => ({
      key: row?.[keyField] || "Unknown",
      outcome: Number(row?.outcome || 0),
      input: Number(row?.input || 0),
      rows: Number(row?.rows || 0),
      efficiency: roles?.costInputMetric ? (Number(row?.input || 0) > 0 ? Number(row?.outcome || 0) / Number(row?.input || 1) : 0) : null,
      lowSignal: isLowSignalKey(row?.[keyField] || "Unknown"),
    }));

    const rawTotalOutcome = normalized.reduce((sum, item) => sum + item.outcome, 0);
    const informative = normalized.filter((item) => !item.lowSignal && item.outcome > 0);
    const informativeTotal = informative.reduce((sum, item) => sum + item.outcome, 0);
    const useInformative = informative.length >= 2 && rawTotalOutcome > 0 && (informativeTotal / rawTotalOutcome) >= 0.2;
    const base = (useInformative ? informative : normalized).sort((a, b) => b.outcome - a.outcome);
    const totalOutcome = base.reduce((sum, item) => sum + item.outcome, 0);
    const topDrivers = base.slice(0, 5).map((item) => ({
      ...item,
      share: totalOutcome > 0 ? item.outcome / totalOutcome : 0,
    }));
    const topShare = topDrivers[0]?.share || 0;
    const secondShare = topDrivers[1]?.share || 0;
    const shareGap = topShare - secondShare;
    const weakSignal = topShare < 0.12 || shareGap < 0.03;
    const weakSignalReason = topShare < 0.12
      ? "top_driver_share_below_threshold"
      : shareGap < 0.03
        ? "top_driver_gap_below_threshold"
        : null;

    return {
      kind: chosen.kind,
      dimensions: chosen.dimensions,
      driverDimensionLabel: chosen.dimensions.join(" | ") || grainChoice?.chosen?.dimension || "entity",
      topDrivers: weakSignal ? [] : topDrivers,
      totalOutcome,
      explanatoryPower: Number(chosen.explanatoryPower || 0),
      filteredLowSignal: useInformative,
      lowSignalShareExcluded: rawTotalOutcome > 0 ? Math.max(0, 1 - (totalOutcome / rawTotalOutcome)) : 0,
      weakSignal,
      weakSignalReason,
      topShare,
      shareGap,
    };
  }

  function detectPatterns(rows, roles, deltas, options = {}) {
    const patterns = {
      momentum: null,
      contribution: null,
      contributionDrivers: null,
      concentrationRisk: null,
      leakage: null,
      watchList: [],
      opportunity: null,
      forecast: null,
      fired: [],
      thresholds: {
        concentrationHigh: 0.6,
        concentrationModerate: 0.4,
        watchCostShare: 0.2,
        watchOutcomeShare: 0.05,
        opportunityEfficiencyLift: 0.2,
        contributionDriverMinShare: 0.12,
        contributionDriverMinGap: 0.03,
        tripleCombinationMaxCardinality: 80,
      },
      grainDiagnostics: null,
    };

    const grainChoice = chooseBestGrain(rows, roles, options);
    const grain = grainChoice?.chosen;
    const contributionDrivers = buildContributionDriverAnalysis(rows, roles, grainChoice, options);
    patterns.grainDiagnostics = {
      chosenGrain: grain?.grain || "single",
      chosenDimension: grain?.dimension || null,
      explanatoryPower: Number(grain?.explanatoryPower || 0),
      comparedSinglePower: Number(grain?.comparedSinglePower || 0),
      capsApplied: grain?.eligibility || null,
      why: grain?.grain === "paired"
        ? "Paired grain explains more incremental outcome within cardinality caps."
        : "Single grain retained due to higher explanatory power or pair cap failure.",
    };

    if (!roles?.primaryOutcomeMetric) {
      return patterns;
    }

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

    const table = (grain?.table || []).map((row) => {
      const key = row?.pairKey || row?.entity || "Unknown";
      const outcome = Number(row?.outcome || 0);
      const input = Number(row?.input || 0);
      const efficiency = roles.costInputMetric ? (input > 0 ? outcome / input : 0) : null;
      const rowsCount = Number(row?.rows || 0);
      return { key, outcome, input, efficiency, rows: rowsCount, lowSignal: isLowSignalKey(key) };
    });

    const rawTotalOutcome = table.reduce((sum, row) => sum + row.outcome, 0);
    const informativeRows = table.filter((row) => !row.lowSignal && row.outcome > 0);
    const informativeOutcome = informativeRows.reduce((sum, row) => sum + row.outcome, 0);
    const shouldUseInformativeOnly = informativeRows.length >= 2 && rawTotalOutcome > 0 && (informativeOutcome / rawTotalOutcome) >= 0.2;
    const rankingBase = shouldUseInformativeOnly ? informativeRows : table;

    const rankedByOutcome = [...rankingBase].sort((a, b) => b.outcome - a.outcome);
    const rankedByOutcomeTop = rankedByOutcome.slice(0, 3);
    const totalOutcome = rankingBase.reduce((sum, row) => sum + row.outcome, 0);
    const totalInput = table.reduce((sum, row) => sum + row.input, 0);

    if (rankedByOutcomeTop.length) {
      patterns.contribution = {
        dimension: grain?.dimension || null,
        topEntities: rankedByOutcomeTop.map((row) => ({ entity: row.key || row.entity || row.pairKey || "Other", value: row.outcome, rawValue: row.outcome })),
        total: totalOutcome,
        filteredLowSignal: shouldUseInformativeOnly,
        lowSignalShareExcluded: rawTotalOutcome > 0 ? Math.max(0, 1 - (totalOutcome / rawTotalOutcome)) : 0,
      };
      const first = rankedByOutcomeTop[0];
      const second = rankedByOutcomeTop[1];
      const topShare = totalOutcome > 0 ? Number(first?.normalizedOutcome || 0) / totalOutcome : 0;
      const top2Share = totalOutcome > 0
        ? (Number(first?.normalizedOutcome || 0) + Number(second?.normalizedOutcome || 0)) / totalOutcome
        : 0;
      patterns.concentrationRisk = {
        topShare,
        top2Share,
        level: topShare > 0.6 ? "High" : top2Share > 0.8 ? "Moderate" : "Low",
      };
      patterns.fired.push("contribution", "concentration");
    }

    if (contributionDrivers?.topDrivers?.length) {
      patterns.contributionDrivers = contributionDrivers;
      patterns.fired.push("cross_dimension_drivers");
    } else if (contributionDrivers?.weakSignal) {
      patterns.contributionDrivers = contributionDrivers;
      patterns.fired.push("cross_dimension_drivers_weak");
    }

    if (roles.costInputMetric && totalInput > 0 && totalOutcome > 0) {
      const watch = table
        .map((row) => ({
          ...row,
          costShare: row.input / totalInput,
          outcomeShare: totalOutcome > 0 ? row.outcome / totalOutcome : 0,
        }))
        .filter((row) => row.costShare > 0.2 && row.outcomeShare < 0.05)
        .sort((a, b) => (b.costShare - b.outcomeShare) - (a.costShare - a.outcomeShare));
      patterns.watchList = watch.slice(0, 3);
      patterns.leakage = patterns.watchList[0] || null;
      if (patterns.watchList.length) patterns.fired.push("leakage");

      const validEff = rankingBase.map((row) => row.efficiency).filter((value) => Number.isFinite(value) && value > 0);
      const med = validEff.length ? median(validEff) : 0;
      const opp = rankingBase
        .filter((row) => Number.isFinite(row.efficiency) && row.efficiency > med * 1.2)
        .sort((a, b) => b.efficiency - a.efficiency)
        .slice(0, 3);
      if (opp.length) {
        patterns.opportunity = { entities: opp, medianEfficiency: med, dimension: grain?.dimension || null };
        patterns.fired.push("opportunity");
      }

      const baseEfficiency = totalInput > 0 ? totalOutcome / totalInput : 0;
      if (baseEfficiency > 0) {
        patterns.forecast = [0.1, 0.2, 0.3].map((increase) => ({
          inputIncrease: increase,
          projectedInput: totalInput * (1 + increase),
          projectedOutcome: totalOutcome + (totalInput * increase * baseEfficiency),
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

    const winners = (patterns?.contribution?.topEntities || []).slice(0, 3);
    const topDriverCombos = (patterns?.contributionDrivers?.topDrivers || []).slice(0, 3);
    const totalOutcome = Number(patterns?.contribution?.total || 0);
    const top = winners[0] || null;
    const runnerUp = winners[1] || null;
    const analysisDimension = patterns?.grainDiagnostics?.chosenDimension
      || patterns?.contribution?.dimension
      || roles?.entityDimensions?.[0]
      || "entity";
    const driverDimensionLabel = patterns?.contributionDrivers?.driverDimensionLabel || analysisDimension;

    const topShare = totalOutcome > 0 && top ? Number(top.value || 0) / totalOutcome : 0;
    const deltaVsRunner = top && runnerUp ? Number(top.value || 0) - Number(runnerUp.value || 0) : 0;
    const deltaShare = totalOutcome > 0 ? deltaVsRunner / totalOutcome : 0;

    if (!roles?.primaryOutcomeMetric) {
      const selectedMetric = roles?.preferredOutcomeMetric || "selected metric";
      return {
        sections: [
          { title: "Executive Summary", bullets: [`What happened: ${selectedMetric} is selected, but numeric coverage is limited for robust ranking.`] },
          { title: "Key Metrics", bullets: ["What succeeded: no statistically stable segment lead can be determined from current parsing quality."] },
          { title: "Next Steps", bullets: ["Why: improve metric parsing/mapping or choose a cleaner numeric metric to unlock ranked insights."] },
        ],
        wordCount: 41,
        context,
      };
    }

    if (totalOutcome <= 0) {
      return {
        sections: [
          { title: "Executive Summary", bullets: [`No measurable activity detected for ${outcome} in the current dataset scope.`] },
          { title: "Key Metrics", bullets: [`${outcome} total is 0. Check selected metric mapping, parsing quality, and missing values.`] },
          { title: "Next Steps", bullets: ["Choose a metric with measurable activity or adjust data quality mappings before drawing conclusions."] },
        ],
        wordCount: 40,
        context,
      };
    }

    const outcomeLabel = outcome;
    const happenedLine = hasTime && deltas?.outcome
      ? `What happened: ${outcomeLabel} moved ${pct(deltas.outcome.pct)} vs prior period to ${compact(deltas.outcome.current || totalOutcome)}.`
      : `What happened: ${outcomeLabel} snapshot is ${compact(totalOutcome)} across ${analysisDimension}.`;
    const pairedMode = String(patterns?.grainDiagnostics?.chosenGrain || "") === "paired";
    const comboLine = pairedMode && winners.length >= 2
      ? ` Top combinations: ${winners[0].entity} and ${winners[1].entity}.`
      : "";
    const crossDriverLine = topDriverCombos.length
      ? `What drove it: ${topDriverCombos[0].key} explains ${pct(topDriverCombos[0].share)} of observed ${outcomeLabel} across ${driverDimensionLabel}.`
      : null;
    const succeededLine = top
      ? `What succeeded: ${top.entity} led ${outcomeLabel} at ${compact(top.value)} (${pct(topShare)} share), +${compact(deltaVsRunner)} vs ${runnerUp?.entity || "next segment"}.${comboLine}`
      : `What succeeded: no stable winning segment identified for ${outcomeLabel}.`;
    const lowSignalNote = patterns?.contribution?.filteredLowSignal
      ? `Low-signal entities were excluded from winner ranking (${pct(patterns.contribution.lowSignalShareExcluded)} of outcome).`
      : null;
    const whyLine = patterns?.concentrationRisk
      ? `Why: ${analysisDimension} concentration is ${patterns.concentrationRisk.level.toLowerCase()} (top share ${pct(patterns.concentrationRisk.topShare)}).`
      : `Why: performance is distributed with no dominant concentration signal.`;
    sections.push({ title: "Executive Summary", bullets: [happenedLine, succeededLine, ...(crossDriverLine ? [crossDriverLine] : []), whyLine, ...(lowSignalNote ? [lowSignalNote] : [])] });

    if (hasTime && patterns?.momentum) {
      const momentumBullets = [`${patterns.momentum}. ${outcome} moved ${pct(deltas?.outcome?.pct)} vs previous period.`];
      if (hasCost && deltas?.cost) momentumBullets.push(`${input} moved ${pct(deltas.cost.pct)} over the same window.`);
      if (hasCost && deltas?.efficiency) momentumBullets.push(`Efficiency (${outcome}/${input}) moved ${pct(deltas.efficiency.pct)}.`);
      sections.push({ title: "Momentum", bullets: momentumBullets });
    } else {
      sections.push({ title: "Momentum", bullets: [`Time trend unavailable. Snapshot based on ${analysisDimension} and ${outcome}.`] });
    }

    const keyBullets = [`Outcome (${outcomeLabel}): ${compact(deltas?.outcome?.current || totalOutcome)}.`];
    if (roles?.preferredOutcomeMetric && roles.preferredOutcomeMetric !== roles.primaryOutcomeMetric) {
      if (roles?.preferredOutcomeReason === "selected_metric_is_input") {
        keyBullets.push(`Selected metric ${roles.preferredOutcomeMetric} behaves like an input metric, so insights use ${roles.primaryOutcomeMetric} as the outcome metric for driver analysis.`);
      } else {
        keyBullets.push(`Selected metric ${roles.preferredOutcomeMetric} is non-additive; insights use ${roles.primaryOutcomeMetric} for valid aggregation.`);
      }
    }
    if (hasCost && deltas?.cost) keyBullets.push(`Input (${input}): ${compact(deltas.cost.current)}.`);
    if (hasCost && deltas?.efficiency) keyBullets.push(`Efficiency (${outcome}/${input}): ${deltas.efficiency.current.toFixed(3)}.`);
    if (deltas?.rate?.metric) keyBullets.push(`${deltas.rate.metric}: ${pct(deltas.rate.current)}.`);
    sections.push({ title: "Key Metrics", bullets: keyBullets });

    sections.push({
      title: "Winners",
      bullets: winners.length
        ? winners.map((item) => {
            const share = totalOutcome > 0 ? Number(item.value || 0) / totalOutcome : 0;
            return `${item.entity}: ${compact(item.value)} ${outcome} (${pct(share)} share).`;
          })
        : [`No stable ${analysisDimension} winners found.`],
    });

    const spotlightBullets = [];
    if (top && runnerUp) {
      spotlightBullets.push(`What succeeded: ${top.entity} outperformed ${runnerUp.entity} by ${compact(deltaVsRunner)} ${outcome}.`);
    }
    if (topDriverCombos.length) {
      const combo = topDriverCombos[0];
      spotlightBullets.push(`Which combinations drove it: ${combo.key} contributed ${compact(combo.outcome)} ${outcome} (${pct(combo.share)} share).`);
    } else if (patterns?.contributionDrivers?.weakSignal) {
      spotlightBullets.push("Which combinations drove it: no single dimension combination cleared the minimum contribution threshold, so the pattern is distributed rather than driver-led.");
    }
    if (patterns?.concentrationRisk) {
      spotlightBullets.push(`Why it happened: ${analysisDimension} concentration is ${patterns.concentrationRisk.level.toLowerCase()} (top share ${pct(patterns.concentrationRisk.topShare)}).`);
    }
    if (hasCost && patterns?.opportunity?.entities?.length) {
      const bestOpp = patterns.opportunity.entities[0];
      spotlightBullets.push(`Efficiency signal: ${bestOpp.key} runs at ${bestOpp.efficiency.toFixed(3)} ${outcome}/${input}.`);
    }
    sections.push({
      title: "Spotlight",
      bullets: spotlightBullets.length ? spotlightBullets : [`No single ${analysisDimension} driver explains the current result.`],
    });

    sections.push({
      title: "Scale Opportunities",
      bullets: hasCost && patterns?.opportunity?.entities?.length
        ? patterns.opportunity.entities.slice(0, 3).map((item) => `${item.key}: ${compact(item.outcome)} ${outcome} on ${compact(item.input)} ${input} (efficiency ${item.efficiency.toFixed(3)}).`)
        : topDriverCombos.length
          ? topDriverCombos.slice(0, 3).map((item) => `${item.key}: ${compact(item.outcome)} ${outcome} (${pct(item.share)} share). Test this combination in adjacent segments before scaling.`)
          : ["No validated efficiency opportunities detected with current input and outcome fields."],
    });

    sections.push({
      title: "Watch List",
      bullets: patterns?.watchList?.length
        ? patterns.watchList.slice(0, 3).map((item) => `${item.key}: ${input} share ${pct(item.costShare)} vs ${outcome} share ${pct(item.outcomeShare)}.`)
        : [`No ${analysisDimension} segments breached watch-list thresholds.`],
    });

    sections.push({
      title: "Forecast Outlook",
      bullets: hasCost && patterns?.forecast
        ? patterns.forecast.map((item) => `+${Math.round(item.inputIncrease * 100)}% ${input} -> projected ${outcome} ${compact(item.projectedOutcome)}.`)
        : ["Forecast unavailable without additive input metric."],
    });

    const next = [];
    if (patterns?.concentrationRisk?.level === "High") next.push(`Protect against over-reliance by expanding beyond ${top?.entity || "the top segment"}.`);
    if (patterns?.opportunity?.entities?.length) next.push(`Prioritize budget toward top efficiency ${analysisDimension} segments first.`);
    if (topDriverCombos.length) next.push(`Validate whether ${topDriverCombos[0].key} remains the top contribution driver in another cut before reallocating budget.`);
    if (patterns?.leakage) next.push(`Audit ${patterns.leakage.key || patterns.leakage.entity || "watch-list segment"} for high-input low-outcome leakage.`);
    sections.push({ title: "Next Steps", bullets: next.length ? next.slice(0, 3) : [`Continue tracking ${analysisDimension} mix and ${outcome} movement each period.`] });

    const words = sections.flatMap((section) => [section.title, ...(section.bullets || [])]).join(" ").split(/\s+/).filter(Boolean);
    if (words.length > 300) {
      sections.forEach((section) => {
        section.bullets = (section.bullets || []).slice(0, 2);
      });
    }
    return { sections, wordCount: words.length, context };
  }

  function generatePerformanceIntelligence(rows, options = {}) {
    const profile = profileDataset(rows || []);
    const roles = detectMetricRoles(profile, options);
    const domain = detectDomain(profile, roles);
    const deltas = computePeriodDeltas(rows || [], roles);
    const patterns = detectPatterns(rows || [], roles, deltas, options);
    const narrative = composeNarrative(roles, patterns, deltas, { domain });
    const narrativeText = (narrative.sections || []).flatMap((section) => [section.title, ...(section.bullets || [])]).join(" ");
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
        grainDiagnostics: patterns.grainDiagnostics || null,
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
    classifyNumericMetric,
    evaluatePairEligibility,
  };

  globalScope.PerformanceInsightsEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
