const fileInput = document.getElementById("fileInput");
const fileInputInline = document.getElementById("fileInputInline");
const loadSample = document.getElementById("loadSample");
const loadSampleInline = document.getElementById("loadSampleInline");
const stateSection = document.getElementById("state");
const errorSection = document.getElementById("error");
const dashboard = document.getElementById("dashboard");
const datasetSummary = document.getElementById("datasetSummary");
const kpiGrid = document.getElementById("kpiGrid");
const lineTitle = document.getElementById("lineTitle");
const lineSubtitle = document.getElementById("lineSubtitle");
const barTitle = document.getElementById("barTitle");
const barSubtitle = document.getElementById("barSubtitle");
const table = document.getElementById("dataTable");

let lineChartInstance = null;
let barChartInstance = null;
let currentTableRows = [];
let currentSort = { key: null, direction: "asc" };

const samplePath = "data sample.csv";

const numberFormatter = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 2,
});

fileInput.addEventListener("change", (event) => handleFile(event.target.files[0]));
fileInputInline.addEventListener("change", (event) => handleFile(event.target.files[0]));
loadSample.addEventListener("click", () => loadSampleData());
loadSampleInline.addEventListener("click", () => loadSampleData());

function handleFile(file) {
  if (!file) return;
  clearError();
  Papa.parse(file, {
    header: true,
    dynamicTyping: false,
    skipEmptyLines: true,
    complete: (results) => {
      if (results.errors.length) {
        showError("We could not parse that CSV. Please check the format and try again.");
        return;
      }
      prepareDashboard(results.data);
    },
  });
}

function loadSampleData() {
  clearError();
  Papa.parse(samplePath, {
    download: true,
    header: true,
    dynamicTyping: false,
    skipEmptyLines: true,
    complete: (results) => {
      if (results.errors.length) {
        showError("Sample CSV could not be loaded. Refresh and try again.");
        return;
      }
      prepareDashboard(results.data);
    },
  });
}

function showError(message) {
  errorSection.textContent = message;
  errorSection.classList.remove("hidden");
}

function clearError() {
  errorSection.classList.add("hidden");
  errorSection.textContent = "";
}

function prepareDashboard(rows) {
  if (!rows || rows.length === 0) {
    showError("No rows found in this CSV.");
    return;
  }

  const columns = Object.keys(rows[0]).filter((key) => key.trim() !== "");
  if (!columns.length) {
    showError("CSV must contain a header row with column names.");
    return;
  }

  const sanitizedRows = rows.map((row) => {
    const sanitized = {};
    for (const key of columns) {
      sanitized[key] = row[key];
    }
    return sanitized;
  });

  const numericColumns = detectNumericColumns(sanitizedRows, columns);
  if (!numericColumns.length) {
    showError("No numeric columns detected. Add at least one numeric metric.");
    return;
  }

  const dateColumn = detectDateColumn(sanitizedRows, columns);
  const categoricalColumns = columns.filter((col) => col !== dateColumn && !numericColumns.includes(col));

  renderKPIs(sanitizedRows, numericColumns);
  renderCharts(sanitizedRows, numericColumns, dateColumn, categoricalColumns);
  renderTable(sanitizedRows, columns);

  datasetSummary.textContent = `${sanitizedRows.length} rows · ${columns.length} columns`;
  stateSection.classList.add("hidden");
  dashboard.classList.remove("hidden");
}

function detectNumericColumns(rows, columns) {
  const metrics = columns.filter((col) => {
    let numericCount = 0;
    let valueCount = 0;
    for (const row of rows) {
      const value = row[col];
      if (value === null || value === undefined || value === "") continue;
      valueCount += 1;
      if (!isNaN(Number(value))) numericCount += 1;
    }
    return valueCount > 0 && numericCount / valueCount >= 0.8;
  });

  return metrics;
}

function detectDateColumn(rows, columns) {
  const candidates = columns.filter((col) => {
    let dateCount = 0;
    let valueCount = 0;
    for (const row of rows) {
      const value = row[col];
      if (!value) continue;
      valueCount += 1;
      const date = new Date(value);
      if (!isNaN(date.valueOf())) dateCount += 1;
    }
    return valueCount > 0 && dateCount / valueCount >= 0.8;
  });

  return candidates[0] || null;
}

function renderKPIs(rows, numericColumns) {
  kpiGrid.innerHTML = "";
  const metrics = numericColumns.slice(0, 6);
  const metricCount = Math.max(3, Math.min(metrics.length, 6));
  const chosen = metrics.slice(0, metricCount);

  chosen.forEach((metric) => {
    const values = rows
      .map((row) => Number(row[metric]))
      .filter((value) => !isNaN(value));
    const total = values.reduce((sum, value) => sum + value, 0);
    const avg = values.length ? total / values.length : 0;

    const card = document.createElement("div");
    card.className = "kpi-card";
    card.innerHTML = `
      <h4>${metric}</h4>
      <div class="kpi-value">${numberFormatter.format(total)}</div>
      <div class="kpi-sub">Avg: ${numberFormatter.format(avg)}</div>
    `;
    kpiGrid.appendChild(card);
  });
}

function renderCharts(rows, numericColumns, dateColumn, categoricalColumns) {
  if (lineChartInstance) lineChartInstance.destroy();
  if (barChartInstance) barChartInstance.destroy();

  const metric = numericColumns[0];
  if (dateColumn) {
    const series = aggregateByDate(rows, dateColumn, metric);
    lineTitle.textContent = `${metric} over time`;
    lineSubtitle.textContent = `Grouped by ${dateColumn}`;
    lineChartInstance = createLineChart(series.labels, series.values);
  } else {
    lineTitle.textContent = "Trend";
    lineSubtitle.textContent = "Date column not detected";
    lineChartInstance = createLineChart([], []);
  }

  const dimension = categoricalColumns[0] || dateColumn || metric;
  const breakdown = aggregateByCategory(rows, dimension, metric);
  barTitle.textContent = `${metric} by ${dimension}`;
  barSubtitle.textContent = "Top categories";
  barChartInstance = createBarChart(breakdown.labels, breakdown.values);
}

function aggregateByDate(rows, dateColumn, metric) {
  const totals = new Map();
  rows.forEach((row) => {
    const dateValue = new Date(row[dateColumn]);
    const metricValue = Number(row[metric]);
    if (isNaN(dateValue.valueOf()) || isNaN(metricValue)) return;
    const key = dateValue.toISOString().slice(0, 10);
    totals.set(key, (totals.get(key) || 0) + metricValue);
  });

  const labels = Array.from(totals.keys()).sort();
  const values = labels.map((label) => totals.get(label));
  return { labels, values };
}

function aggregateByCategory(rows, dimension, metric) {
  const totals = new Map();
  rows.forEach((row) => {
    const key = row[dimension] ? String(row[dimension]) : "Unknown";
    const metricValue = Number(row[metric]);
    if (isNaN(metricValue)) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
  });

  const sorted = Array.from(totals.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  return {
    labels: sorted.map(([label]) => label),
    values: sorted.map(([, value]) => value),
  };
}

function createLineChart(labels, values) {
  const ctx = document.getElementById("lineChart");
  return new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: values,
          borderColor: "#2f5e4e",
          backgroundColor: "rgba(47, 94, 78, 0.15)",
          tension: 0.35,
          fill: true,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: { beginAtZero: true },
      },
    },
  });
}

function createBarChart(labels, values) {
  const ctx = document.getElementById("barChart");
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: values,
          backgroundColor: "#f2c14e",
          borderRadius: 8,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: { beginAtZero: true },
      },
    },
  });
}

function renderTable(rows, columns) {
  currentTableRows = rows.slice(0, 50);
  currentSort = { key: null, direction: "asc" };

  table.innerHTML = "";
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.addEventListener("click", () => sortTable(col));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  currentTableRows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((col) => {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
}

function sortTable(column) {
  const direction = currentSort.key === column && currentSort.direction === "asc" ? "desc" : "asc";
  currentSort = { key: column, direction };

  const sorted = [...currentTableRows].sort((a, b) => {
    const aVal = a[column] ?? "";
    const bVal = b[column] ?? "";
    const aNum = Number(aVal);
    const bNum = Number(bVal);
    if (!isNaN(aNum) && !isNaN(bNum)) {
      return direction === "asc" ? aNum - bNum : bNum - aNum;
    }
    return direction === "asc"
      ? String(aVal).localeCompare(String(bVal))
      : String(bVal).localeCompare(String(aVal));
  });

  currentTableRows = sorted;
  renderTable(currentTableRows, Object.keys(currentTableRows[0] || {}));
}
