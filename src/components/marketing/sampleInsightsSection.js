import { getSampleInsights } from "../../lib/sampleFeed.js";

const SUPPORT_THRESHOLD_TOTAL = 30;
const SUPPORT_THRESHOLD_GROUP = 5;

function createSkeletonCard() {
  const card = document.createElement("div");
  card.className = "insight-card skeleton-card";
  card.innerHTML = `
    <div class="skeleton-line wide"></div>
    <div class="skeleton-line"></div>
    <div class="skeleton-line"></div>
  `;
  return card;
}

function createEmptyState(onCtaClick) {
  const wrap = document.createElement("div");
  wrap.className = "empty-state";
  wrap.innerHTML = `
    <p>No sample insights are available yet.</p>
    <button class="primary">Start free analysis</button>
  `;
  const button = wrap.querySelector("button");
  if (button && onCtaClick) button.addEventListener("click", onCtaClick);
  return wrap;
}

function createInsightCard(insight) {
  const card = document.createElement("div");
  card.className = "insight-card";
  const currentValue = insight.metric.format(insight.value.currentValue);
  const deltaLabel = insight.value.deltaType === "percentagePoints"
    ? `${(insight.value.deltaAbsolute * 100).toFixed(1)} pp`
    : `${(insight.value.deltaRelative * 100).toFixed(1)}%`;
  const windowLabel = `${insight.period.currentPeriodStart} → ${insight.period.currentPeriodEnd} vs ${insight.period.comparisonPeriodStart} → ${insight.period.comparisonPeriodEnd} (${insight.period.comparisonType})`;
  const supportTotal = insight.evidence.sampleSize?.total || 0;
  const supportGroups = Object.values(insight.evidence.sampleSize?.byGroup || {});
  const minGroup = supportGroups.length ? Math.min(...supportGroups) : 0;
  const hasSupport = supportTotal >= SUPPORT_THRESHOLD_TOTAL && minGroup >= SUPPORT_THRESHOLD_GROUP;
  const supportText = hasSupport ? `Support: ${insight.confidence.supportLevel}` : "Support: low";
  const limitation = hasSupport ? "" : "Insufficient data to make a definitive claim.";

  card.innerHTML = `
    <div class="insight-header">
      <span class="severity ${insight.category.toLowerCase()}">${insight.category}</span>
    </div>
    <h4>${insight.title}</h4>
    <p class="insight-meta">${windowLabel}</p>
    <div class="insight-metric">
      <span>${insight.metric.name}</span>
      <strong title="${insight.metric.definition}">${currentValue}</strong>
      <em>${deltaLabel}</em>
    </div>
    <p class="insight-meta">Evidence: ${insight.evidence.primaryChartType} · ${insight.evidence.breakdownDimension} · n=${supportTotal}</p>
    <p class="insight-meta">${supportText}${limitation ? ` · ${limitation}` : ""}</p>
    <div class="insight-actions">
      ${insight.actions.slice(0, 3).map((action) => `<div>• ${action.text}</div>`).join("")}
    </div>
    <button class="ghost view-evidence" data-id="${insight.id}">View evidence</button>
  `;
  return card;
}

export async function renderSampleInsightsSection(container) {
  if (!container) return;
  container.innerHTML = "";
  container.appendChild(createSkeletonCard());
  container.appendChild(createSkeletonCard());
  container.appendChild(createSkeletonCard());

  const ctaTarget = document.getElementById("startFreeBtn");
  const insights = await getSampleInsights();

  container.innerHTML = "";
  if (!insights || insights.length === 0) {
    container.appendChild(createEmptyState(() => ctaTarget?.scrollIntoView({ behavior: "smooth" })));
    return;
  }

  const drilldown = document.getElementById("insightDrilldown");
  insights.forEach((insight) => {
    const card = createInsightCard(insight);
    const button = card.querySelector(".view-evidence");
    if (button && drilldown) {
      button.addEventListener("click", () => {
        renderDrilldown(drilldown, insight);
      });
    }
    container.appendChild(card);
  });
}

function renderDrilldown(container, insight) {
  if (!container) return;
  container.classList.remove("hidden");
  const groups = insight.drilldown?.topSegments || [];
  const rows = insight.drilldown?.rows || [];
  const chart = buildMiniChart(insight.evidence);

  container.innerHTML = `
    <div class="drilldown-header">
      <h4>Evidence: ${insight.title}</h4>
      <span class="helper-text">${insight.metric.definition}</span>
    </div>
    <div class="drilldown-chart">${chart}</div>
    <div class="drilldown-body">
      <div>
        <strong>Top segments</strong>
        <div>${groups.map((item) => `<div>• ${item}</div>`).join("") || "—"}</div>
      </div>
      <div>
        <strong>Rows preview</strong>
        <div class="drilldown-rows">
          ${rows.map((row) => `<div>${Object.entries(row).map(([k, v]) => `${k}: ${v}`).join(" · ")}</div>`).join("")}
        </div>
      </div>
    </div>
  `;
}

function buildMiniChart(evidence) {
  const points = evidence.chartPoints || [];
  if (!points.length) return "<div class=\"helper-text\">No chart data</div>";
  const width = 320;
  const height = 140;
  const padding = 10;
  const xs = points.map((p) => new Date(p.x).getTime());
  const ys = points.map((p) => p.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const scaleX = (x) => padding + ((x - minX) / (maxX - minX || 1)) * (width - padding * 2);
  const scaleY = (y) => height - padding - ((y - minY) / (maxY - minY || 1)) * (height - padding * 2);

  const grouped = points.reduce((acc, point) => {
    const key = point.group || "Total";
    acc[key] = acc[key] || [];
    acc[key].push(point);
    return acc;
  }, {});

  const paths = Object.entries(grouped).map(([key, groupPoints]) => {
    const path = groupPoints
      .map((p, idx) => `${idx === 0 ? "M" : "L"}${scaleX(new Date(p.x).getTime())},${scaleY(p.y)}`)
      .join(" ");
    return `<path d="${path}" fill="none" stroke="#a78bfa" stroke-width="2" />`;
  }).join("");

  return `
    <svg class="drilldown-sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      ${paths}
    </svg>
  `;
}
