import { getSampleInsights } from "../../lib/sampleFeed.js";

function severityLabel(severity) {
  if (severity === "growth") return "Growth";
  if (severity === "risk") return "Risk";
  return "Opportunity";
}

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
  card.innerHTML = `
    <div class="insight-header">
      <span class="severity ${insight.severity}">${severityLabel(insight.severity)}</span>
    </div>
    <h4>${insight.title}</h4>
    <p>${insight.subtitle}</p>
    <div class="insight-metric">
      <span>${insight.metricLabel}</span>
      <strong>${insight.metricValue}</strong>
      <em>${insight.metricDelta}</em>
    </div>
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

  insights.forEach((insight) => {
    container.appendChild(createInsightCard(insight));
  });
}
