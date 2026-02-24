import { getMiniDashboard } from "../../lib/sampleFeed.js";

function buildSparkline(points) {
  if (!points.length) return "";
  const width = 260;
  const height = 120;
  const padding = 8;
  const xs = points.map((p) => p.x);
  const ys = points.map((p) => p.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);

  const scaleX = (x) => {
    if (maxX === minX) return padding;
    return padding + ((x - minX) / (maxX - minX)) * (width - padding * 2);
  };
  const scaleY = (y) => {
    if (maxY === minY) return height / 2;
    return height - padding - ((y - minY) / (maxY - minY)) * (height - padding * 2);
  };

  const path = points
    .map((point, idx) => `${idx === 0 ? "M" : "L"}${scaleX(point.x)},${scaleY(point.y)}`)
    .join(" ");

  return `
    <svg class="sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-hidden="true">
      <defs>
        <linearGradient id="sparklineFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#a78bfa" stop-opacity="0.4" />
          <stop offset="100%" stop-color="#a78bfa" stop-opacity="0" />
        </linearGradient>
      </defs>
      <path d="${path}" fill="none" stroke="#a78bfa" stroke-width="2.5" />
    </svg>
  `;
}

function createSkeleton() {
  const wrap = document.createElement("div");
  wrap.className = "mini-dashboard";
  wrap.innerHTML = `
    <div class="mini-kpis">
      <div class="kpi-pill skeleton-line wide"></div>
      <div class="kpi-pill skeleton-line wide"></div>
      <div class="kpi-pill skeleton-line wide"></div>
    </div>
    <div class="skeleton-line wide" style="height:120px"></div>
  `;
  return wrap;
}

function createEmptyState() {
  const wrap = document.createElement("div");
  wrap.className = "empty-state";
  wrap.innerHTML = "<p>Mini dashboard preview unavailable.</p>";
  return wrap;
}

export async function renderMiniDashboardPreview(container) {
  if (!container) return;
  container.classList.add("mini-dashboard");
  container.innerHTML = "";
  container.appendChild(createSkeleton());

  const payload = await getMiniDashboard();
  container.innerHTML = "";
  if (!payload || !payload.miniKpis || payload.miniKpis.length === 0) {
    container.appendChild(createEmptyState());
    return;
  }

  const kpiRow = document.createElement("div");
  kpiRow.className = "mini-kpis";
  payload.miniKpis.forEach((kpi) => {
    const pill = document.createElement("div");
    pill.className = "kpi-pill";
    pill.innerHTML = `
      <span>${kpi.label}</span>
      <strong title="${kpi.definition}">${kpi.value}</strong>
      <em>${kpi.delta}</em>
      <small class="kpi-caption">${kpi.window}</small>
    `;
    kpiRow.appendChild(pill);
  });

  const chart = document.createElement("div");
  chart.className = "mini-chart";
  chart.innerHTML = buildSparkline(payload.miniTrend || []);

  container.appendChild(kpiRow);
  container.appendChild(chart);
}
