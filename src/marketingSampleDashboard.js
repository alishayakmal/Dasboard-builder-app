import {
  loadSampleDataset,
  getMetricConfig,
  computeKpis,
  computeKpiMiniTrend,
  computeTrend,
  computeBreakdown,
  computeInsights,
  formatValue,
} from "./lib/sampleDashboard.js";

const state = {
  rows: [],
  metric: "revenue",
  breakdown: "overall",
  dimension: "plan",
  evidenceInsightId: null,
};

export async function renderSampleDashboard() {
  const kpiRoot = document.getElementById("sampleKpiGrid");
  const trendMetric = document.getElementById("sampleTrendMetric");
  const trendBreakdown = document.getElementById("sampleTrendBreakdown");
  const trendChart = document.getElementById("sampleTrendChart");
  const dimensionSelect = document.getElementById("sampleDimensionSelect");
  const breakdownChart = document.getElementById("sampleBreakdownChart");
  const evidencePanel = document.getElementById("sampleEvidencePanel");
  const insightsRoot = document.getElementById("sampleInsightsRoot");

  if (!kpiRoot || !trendMetric || !trendBreakdown || !trendChart || !dimensionSelect || !breakdownChart || !insightsRoot) {
    console.warn("Sample dashboard elements missing. Retrying...");
    setTimeout(() => renderSampleDashboard(), 150);
    return;
  }

  const metrics = getMetricConfig();
  metrics.forEach((metric) => {
    const option = document.createElement("option");
    option.value = metric.key;
    option.textContent = metric.label;
    trendMetric.appendChild(option);
  });
  trendMetric.value = state.metric;

  trendMetric.addEventListener("change", () => {
    state.metric = trendMetric.value;
    renderTrend(trendChart);
    renderBreakdown(breakdownChart);
    renderInsights(insightsRoot, evidencePanel, breakdownChart);
  });

  trendBreakdown.addEventListener("change", () => {
    state.breakdown = trendBreakdown.value;
    renderTrend(trendChart);
  });

  dimensionSelect.addEventListener("change", () => {
    state.dimension = dimensionSelect.value;
    renderBreakdown(breakdownChart);
    renderInsights(insightsRoot, evidencePanel, breakdownChart);
  });

  kpiRoot.innerHTML = "<div class=\"skeleton-line wide\"></div>";
  trendChart.innerHTML = "<div class=\"skeleton-line wide\"></div>";
  breakdownChart.innerHTML = "<div class=\"skeleton-line wide\"></div>";
  insightsRoot.innerHTML = "<div class=\"skeleton-line wide\"></div>";

  state.rows = await loadSampleDataset();
  renderKpis(kpiRoot);
  renderTrend(trendChart);
  renderBreakdown(breakdownChart);
  renderInsights(insightsRoot, evidencePanel, breakdownChart);
}

function renderKpis(container) {
  const kpis = computeKpis(state.rows);
  container.innerHTML = "";
  kpis.forEach((kpi) => {
    const card = document.createElement("div");
    card.className = "kpi-pill";
    const deltaLabel = formatKpiDelta(kpi);
    const deltaArrow = kpi.direction === "up" ? "▲" : kpi.direction === "down" ? "▼" : "•";
    const metricConfig = getMetricConfig().find((m) => m.key === kpi.key);
    const miniTrend = computeKpiMiniTrend(state.rows, kpi.key);
    card.innerHTML = `
      <span>${kpi.label}</span>
      <strong title="${kpi.definition}">${formatValue(kpi.value, metricConfig.unit)}</strong>
      <em class="kpi-delta ${kpi.direction}">${deltaArrow} ${deltaLabel}</em>
      <div class="kpi-mini-trend">${buildKpiMiniTrendSvg(miniTrend, metricConfig.unit, kpi.label)}</div>
      <small class="kpi-caption">${kpi.window}</small>
    `;
    container.appendChild(card);
  });
}

function formatKpiDelta(kpi) {
  if (kpi.deltaType === "percentagePoints") {
    const signed = (kpi.delta * 100).toFixed(1);
    return `${Number(signed) > 0 ? "+" : Number(signed) < 0 ? "" : ""}${signed} pp`;
  }
  const signed = (kpi.delta * 100).toFixed(1);
  return `${Number(signed) > 0 ? "+" : Number(signed) < 0 ? "" : ""}${signed}%`;
}

function buildKpiMiniTrendSvg(series, unit, label) {
  const currentSeries = Array.isArray(series?.currentSeries) ? series.currentSeries : [];
  const priorSeries = Array.isArray(series?.priorSeries) ? series.priorSeries : [];
  const length = Math.max(currentSeries.length, priorSeries.length, 30);
  const width = 240;
  const height = 50;
  const padX = 4;
  const padY = 4;
  const values = [...currentSeries, ...priorSeries];
  const maxY = Math.max(...values, 0);
  const minY = Math.min(...values, 0);
  const scaleX = (index) => padX + (index / Math.max(length - 1, 1)) * (width - padX * 2);
  const scaleY = (value) => (height - padY) - ((value - minY) / (Math.max(maxY - minY, 1))) * (height - padY * 2);
  const linePath = (points) => points.map((value, index) => `${index === 0 ? "M" : "L"}${scaleX(index)},${scaleY(value)}`).join(" ");
  const currentPath = linePath(currentSeries);
  const priorPath = linePath(priorSeries);
  const lastCurrent = currentSeries[currentSeries.length - 1] ?? 0;
  const lastPrior = priorSeries[priorSeries.length - 1] ?? 0;

  return `
    <svg class="kpi-sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-label="${label} 30 day comparison sparkline">
      <path d="${priorPath}" fill="none" stroke="rgba(161,161,170,0.65)" stroke-width="1.5">
        <title>Prior 30 days (day 30): ${formatValue(lastPrior, unit)}</title>
      </path>
      <path d="${currentPath}" fill="none" stroke="rgba(52,211,153,0.9)" stroke-width="1.8">
        <title>Current 30 days (day 30): ${formatValue(lastCurrent, unit)}</title>
      </path>
    </svg>
  `;
}

function renderTrend(container) {
  const metricConfig = getMetricConfig().find((m) => m.key === state.metric) || { label: state.metric, unit: "count" };
  const series = computeTrend(state.rows, state.metric, state.breakdown);
  warnIfIdenticalTrendSeries(series, state.breakdown);
  container.innerHTML = buildLineChart(series, {
    metricLabel: metricConfig.label,
    unit: metricConfig.unit,
    breakdown: state.breakdown,
  });
}

function renderBreakdown(container) {
  const breakdown = computeBreakdown(state.rows, state.metric, state.dimension);
  container.innerHTML = buildBarChart(breakdown);
}

function renderInsights(container, evidencePanel, breakdownChart) {
  console.log("INSIGHT INPUT", state.metric, state.dimension);
  const insights = computeInsights(state.rows, state.metric, state.dimension);
  container.innerHTML = "";
  insights.forEach((insight) => {
    const confidence = insight.confidence && typeof insight.confidence === "object"
      ? insight.confidence
      : { level: (insight.confidenceLevel || insight.confidence || "low"), reasons: insight.confidenceReasons || [], metrics: insight.confidenceDetails || {} };
    const confidenceLevel = confidence.level || "low";
    const confidenceLabel = `${confidenceLevel.charAt(0).toUpperCase()}${confidenceLevel.slice(1)}`;
    const whyLine = confidenceLevel !== "high" && Array.isArray(confidence.reasons) && confidence.reasons.length
      ? `<p class="insight-meta confidence-why"><strong>Why this confidence:</strong> ${confidence.reasons.join(" ")}</p>`
      : "";
    const card = document.createElement("div");
    card.className = "insight-card";
    card.innerHTML = `
      <h4>${insight.headline || insight.claim}</h4>
      <p class="insight-meta">${insight.deltaSummary || insight.evidenceSummary || insight.evidence || ""}</p>
      <p class="insight-meta">Driver: ${insight.driver}</p>
      <p class="insight-meta">Action: ${insight.actionSummary || insight.action}</p>
      <span
        class="severity ${confidenceLevel} confidence-pill"
        tabindex="0"
        aria-label="Confidence info: Confidence reflects how reliable this insight is based on sample size, size of the difference and consistency over time."
      >
        Confidence: ${confidenceLabel}
        <span class="confidence-tooltip" role="tooltip">Confidence reflects how reliable this insight is based on sample size, size of the difference and consistency over time.</span>
      </span>
      ${whyLine}
      <button type="button" class="ghost view-evidence" data-insight-id="${insight.id}">View evidence</button>
    `;
    const button = card.querySelector(".view-evidence");
    if (button) {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        state.evidenceInsightId = insight.id;
        console.log("view-evidence", insight.id);
        showEvidence(evidencePanel, insight);
        breakdownChart?.scrollIntoView({ behavior: "smooth" });
      });
    }
    container.appendChild(card);
  });

  if (evidencePanel && !evidencePanel.classList.contains("hidden")) {
    const activeInsight = insights.find((insight) => insight.id === state.evidenceInsightId) || insights[0];
    if (activeInsight) {
      showEvidence(evidencePanel, activeInsight);
      state.evidenceInsightId = activeInsight.id;
    } else {
      evidencePanel.classList.add("hidden");
      evidencePanel.innerHTML = "";
      state.evidenceInsightId = null;
    }
  }
}

function showEvidence(panel, insight) {
  if (!panel) return;
  panel.classList.remove("hidden");
  const evidence = insight.evidence || {};
  const confidence = insight.confidence && typeof insight.confidence === "object"
    ? insight.confidence
    : { level: (insight.confidenceLevel || insight.confidence || "low"), reasons: insight.confidenceReasons || [], metrics: insight.confidenceDetails || {} };
  const comparisonRows = Array.isArray(evidence.comparisonTable) ? evidence.comparisonTable : [];
  const comparisonTable = comparisonRows.length
    ? `
      <div class="evidence-table-wrap">
        <table class="evidence-table">
          <thead>
            <tr>
              <th>Segment</th>
              <th>Value</th>
              <th>Share</th>
              <th>Rank</th>
            </tr>
          </thead>
          <tbody>
            ${comparisonRows.map((row) => `
              <tr>
                <td>${row.segment}</td>
                <td>${row.valueLabel ?? row.value}</td>
                <td>${row.shareLabel ?? ""}</td>
                <td>${row.rank}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `
    : "<div class=\"helper-text\">No comparison evidence available.</div>";
  const trendSummary = evidence.trendSummary || {};
  const periodsAnalysed = Number.isFinite(confidence.metrics?.periods) ? confidence.metrics.periods : (trendSummary.periodsAnalysed ?? 0);
  const periodsExplanation = periodsAnalysed === 1
    ? "Only one time period available. Trend reliability is limited."
    : periodsAnalysed > 1 && periodsAnalysed < 3
      ? `Limited history: this insight is based on only ${periodsAnalysed} time period(s).`
      : `Based on ${periodsAnalysed} time periods.`;
  const confidenceReasons = Array.isArray(confidence.reasons) && confidence.reasons.length
    ? confidence.reasons.map((reason) => `<li>${reason}</li>`).join("")
    : "<li>No major reliability issues detected.</li>";
  panel.innerHTML = `
    <div class="evidence-section">
      <strong>Segment comparison</strong>
      ${comparisonTable}
    </div>
    <div class="evidence-section">
      <strong>What drives this insight</strong>
      <div>${evidence.contributionSummary || "No contribution summary available."}</div>
    </div>
    <div class="evidence-section">
      <strong>Trend consistency</strong>
      <div class="evidence-stats">
        <div class="evidence-stat">
          <span class="inline-info">
            Periods analysed
            <span
              class="info-dot"
              tabindex="0"
              aria-label="Periods analysed tooltip: Periods analysed represents how many time buckets, such as months or days, were used to validate this insight over time."
            >
              i
              <span class="info-tooltip" role="tooltip">Periods analysed represents how many time buckets, such as months or days, were used to validate this insight over time.</span>
            </span>
          </span>
          <strong>${periodsAnalysed}</strong>
        </div>
        <div class="evidence-stat"><span>Stability score</span><strong>${evidence.consistencyScore ?? 0}</strong></div>
        <div class="evidence-stat"><span>Variance</span><strong>${Number.isFinite(trendSummary.variance) ? trendSummary.variance.toFixed(2) : "0.00"}</strong></div>
      </div>
      <div class="helper-text">${periodsExplanation}</div>
      <div class="helper-text">${(trendSummary.leadPeriods ?? 0)} of ${(trendSummary.comparedPeriods ?? 0)} overlapping periods favor the top segment.</div>
    </div>
    <div class="evidence-section">
      <strong>Confidence explanation</strong>
      <ul class="evidence-reasons">${confidenceReasons}</ul>
    </div>
  `;
}

function buildLineChart(series, options = {}) {
  const metricLabel = options.metricLabel || "Metric";
  const unit = options.unit || "count";
  const breakdown = options.breakdown || "overall";
  const width = 640;
  const height = 260;
  const paddingX = 24;
  const paddingTop = 16;
  const paddingBottom = 42;
  const allPoints = series.flatMap((s) => s.points);
  if (!allPoints.length) return "<div class=\"helper-text\">No data</div>";
  const xs = allPoints.map((p) => p.monthTs);
  const ys = allPoints.map((p) => p.value);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);

  const scaleX = (x) => paddingX + ((x - minX) / (maxX - minX || 1)) * (width - paddingX * 2);
  const scaleY = (y) => (height - paddingBottom) - ((y - minY) / (maxY - minY || 1)) * ((height - paddingBottom) - paddingTop);
  const colors = ["#60a5fa", "#34d399", "#f59e0b", "#f472b6", "#a78bfa", "#fb7185"];
  const monthTicks = Array.from(new Set(allPoints.map((point) => point.monthTs))).sort((a, b) => a - b);

  const paths = series.map((s, seriesIndex) => {
    const color = colors[seriesIndex % colors.length];
    const path = s.points
      .map((p, idx) => `${idx === 0 ? "M" : "L"}${scaleX(p.monthTs)},${scaleY(p.value)}`)
      .join(" ");
    return `<path d="${path}" fill="none" stroke="${color}" stroke-width="2" />`;
  }).join("");

  const dots = series.map((s, seriesIndex) => s.points.map((p) => {
    const color = colors[seriesIndex % colors.length];
    const label = new Date(p.monthTs).toLocaleString("en-US", { month: "short", year: "numeric" });
    const cx = scaleX(p.monthTs);
    const cy = scaleY(p.value);
    const tooltip = `${label} · ${s.group}: ${formatValue(p.value, unit)} (${metricLabel})`;
    return `
      <g>
        <circle cx="${cx}" cy="${cy}" r="6" fill="transparent">
          <title>${tooltip}</title>
        </circle>
        <circle cx="${cx}" cy="${cy}" r="3" fill="${color}" pointer-events="none"></circle>
      </g>
    `;
  }).join("")).join("");

  const xAxisLineY = height - paddingBottom;
  const xAxis = `<line x1="${paddingX}" y1="${xAxisLineY}" x2="${width - paddingX}" y2="${xAxisLineY}" stroke="rgba(255,255,255,0.18)" stroke-width="1" />`;
  const xTicks = monthTicks.map((monthTs) => {
    const x = scaleX(monthTs);
    const label = new Date(monthTs).toLocaleString("en-US", { month: "short" });
    const fullLabel = new Date(monthTs).toLocaleString("en-US", { month: "short", year: "numeric" });
    return `
      <g>
        <line x1="${x}" y1="${xAxisLineY}" x2="${x}" y2="${xAxisLineY + 4}" stroke="rgba(255,255,255,0.14)" />
        <text x="${x}" y="${height - 18}" text-anchor="middle" fill="rgba(255,255,255,0.65)" font-size="10">
          ${label}
          <title>${fullLabel}</title>
        </text>
      </g>
    `;
  }).join("");

  const legend = series.map((s, seriesIndex) => {
    const color = colors[seriesIndex % colors.length];
    return `
      <span class="trend-legend-item">
        <span class="trend-legend-swatch" style="background:${color}"></span>
        <span>${s.group}</span>
      </span>
    `;
  }).join("");

  const breakdownLabel = breakdown === "overall" ? "Overall" : `By ${breakdown}`;
  const debug = `
    <div class="trend-debug helper-text">
      Metric: ${metricLabel} · Breakdown: ${breakdownLabel} · Datasets: ${series.length}
      <br />
      Dataset labels: ${series.map((s) => s.group).join(", ")}
    </div>
  `;

  return `
    <div class="trend-legend">${legend}</div>
    <svg class="trend-line" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      ${xAxis}
      ${xTicks}
      ${paths}
      ${dots}
    </svg>
    ${debug}
  `;
}

function warnIfIdenticalTrendSeries(series, breakdown) {
  if (breakdown === "overall" || series.length < 2) return;
  const baseline = series[0]?.points?.map((p) => p.value).join("|");
  if (!baseline) return;
  const allIdentical = series.every((entry) => entry.points.map((p) => p.value).join("|") === baseline);
  if (allIdentical) {
    console.warn("Trend breakdown series identical, check breakdown field mapping", breakdown);
  }
}

function buildBarChart(items) {
  if (!items.length) return "<div class=\"helper-text\">No data</div>";
  const maxValue = Math.max(...items.map((item) => item.value));
  return `
    <div class="bar-chart">
      ${items.map((item) => {
        const width = (item.value / maxValue) * 100;
        return `
          <div class="bar-row">
            <span>${item.key}</span>
            <div class="bar">
              <div class="bar-fill" style="width:${width}%"></div>
            </div>
            <strong>${formatValue(item.value, getMetricConfig().find((m) => m.key === state.metric).unit)}</strong>
          </div>
        `;
      }).join("")}
    </div>
  `;
}
