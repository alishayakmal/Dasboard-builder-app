import {
  loadSampleDataset,
  getMetricConfig,
  computeKpis,
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
    renderInsights(insightsRoot);
  });

  trendBreakdown.addEventListener("change", () => {
    state.breakdown = trendBreakdown.value;
    renderTrend(trendChart);
  });

  dimensionSelect.addEventListener("change", () => {
    state.dimension = dimensionSelect.value;
    renderBreakdown(breakdownChart);
    renderInsights(insightsRoot);
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
    const deltaLabel = kpi.deltaType === "percentagePoints"
      ? `${(kpi.delta * 100).toFixed(1)} pp`
      : `${(kpi.delta * 100).toFixed(1)}%`;
    const metricConfig = getMetricConfig().find((m) => m.key === kpi.key);
    card.innerHTML = `
      <span>${kpi.label}</span>
      <strong title="${kpi.definition}">${formatValue(kpi.value, metricConfig.unit)}</strong>
      <em>${deltaLabel}</em>
      <small class="kpi-caption">${kpi.window}</small>
    `;
    container.appendChild(card);
  });
}

function renderTrend(container) {
  const series = computeTrend(state.rows, state.metric, state.breakdown);
  container.innerHTML = buildLineChart(series);
}

function renderBreakdown(container) {
  const breakdown = computeBreakdown(state.rows, state.metric, state.dimension);
  container.innerHTML = buildBarChart(breakdown);
}

function renderInsights(container, evidencePanel, breakdownChart) {
  const insights = computeInsights(state.rows, state.metric, state.dimension);
  container.innerHTML = "";
  insights.forEach((insight) => {
    const card = document.createElement("div");
    card.className = "insight-card";
    card.innerHTML = `
      <h4>${insight.claim}</h4>
      <p class="insight-meta">${insight.evidence}</p>
      <p class="insight-meta">Driver: ${insight.driver}</p>
      <p class="insight-meta">Action: ${insight.action}</p>
      <span class="severity ${insight.confidence}">Confidence: ${insight.confidence}</span>
      <button class="ghost view-evidence">View evidence</button>
    `;
    const button = card.querySelector(".view-evidence");
    if (button) {
      button.addEventListener("click", () => {
        console.log("View evidence clicked", insight);
        showEvidence(evidencePanel, insight);
        breakdownChart?.scrollIntoView({ behavior: "smooth" });
      });
    }
    container.appendChild(card);
  });
}

function showEvidence(panel, insight) {
  if (!panel) return;
  panel.classList.remove("hidden");
  panel.innerHTML = `
    <strong>Evidence detail</strong>
    <div>${insight.evidence}</div>
    <div>${insight.driver}</div>
    <div>${insight.action}</div>
    <div>Confidence: ${insight.confidence}</div>
  `;
}

function buildLineChart(series) {
  const width = 640;
  const height = 220;
  const padding = 20;
  const allPoints = series.flatMap((s) => s.points);
  if (!allPoints.length) return "<div class=\"helper-text\">No data</div>";
  const xs = allPoints.map((p) => p.monthTs);
  const ys = allPoints.map((p) => p.value);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);

  const scaleX = (x) => padding + ((x - minX) / (maxX - minX || 1)) * (width - padding * 2);
  const scaleY = (y) => height - padding - ((y - minY) / (maxY - minY || 1)) * (height - padding * 2);

  const paths = series.map((s) => {
    const path = s.points
      .map((p, idx) => `${idx === 0 ? "M" : "L"}${scaleX(p.monthTs)},${scaleY(p.value)}`)
      .join(" ");
    return `<path d="${path}" fill="none" stroke="#a78bfa" stroke-width="2" />`;
  }).join("");

  const dots = series.map((s) => s.points.map((p) => {
    const label = new Date(p.monthTs).toLocaleString("en-US", { month: "short", year: "numeric" });
    return `<circle cx="${scaleX(p.monthTs)}" cy="${scaleY(p.value)}" r="3" fill="#a78bfa">
      <title>${label}: ${p.value.toFixed(0)}</title>
    </circle>`;
  }).join("")).join("");

  return `
    <svg class="trend-line" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      ${paths}
      ${dots}
    </svg>
  `;
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
