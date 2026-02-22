const fileInput = document.getElementById("fileInput");
const fileInputInline = document.getElementById("fileInputInline");
const loadSample = document.getElementById("loadSample");
const loadSampleInline = document.getElementById("loadSampleInline");
const stateSection = document.getElementById("state");
const errorSection = document.getElementById("error");
const warningsSection = document.getElementById("warnings");
const statusSection = document.getElementById("status");
const dashboard = document.getElementById("dashboard");
const datasetSummary = document.getElementById("datasetSummary");
const kpiGrid = document.getElementById("kpiGrid");
const table = document.getElementById("dataTable");
const metricSelect = document.getElementById("metricSelect");
const dimensionSelect = document.getElementById("dimensionSelect");
const topNSelect = document.getElementById("topN");
const domainSelect = document.getElementById("domainSelect");
const dateStartInput = document.getElementById("dateStart");
const dateEndInput = document.getElementById("dateEnd");
const exportCsvButton = document.getElementById("exportCsv");
const downloadInsightsButton = document.getElementById("downloadInsights");
const trendTitle = document.getElementById("trendTitle");
const trendSubtitle = document.getElementById("trendSubtitle");
const barTitle = document.getElementById("barTitle");
const barSubtitle = document.getElementById("barSubtitle");
const extraTitle = document.getElementById("extraTitle");
const extraSubtitle = document.getElementById("extraSubtitle");
const insightsList = document.getElementById("insightsList");
const profileTable = document.getElementById("profileTable");
const qualityBadge = document.getElementById("qualityBadge");
const buildStamp = document.getElementById("buildStamp");
const tabButtons = document.querySelectorAll(".tab-button");
const tabPanels = document.querySelectorAll("[data-tab-panel]");
const dropZone = document.getElementById("dropZone");
const uploadTrigger = document.getElementById("uploadTrigger");
const uploadInput = document.getElementById("uploadInput");
const apiBaseUrl = document.getElementById("apiBaseUrl");
const apiAuthType = document.getElementById("apiAuthType");
const apiKey = document.getElementById("apiKey");
const apiEndpoint = document.getElementById("apiEndpoint");
const testApiButton = document.getElementById("testApi");
const apiStatus = document.getElementById("apiStatus");
const sampleGallery = document.getElementById("sampleGallery");
const suggestedTrends = document.getElementById("suggestedTrends");

window.addEventListener("error", (event) => {
  const detail = event?.message ? `Details: ${event.message}` : "Details: unknown error";
  console.error("Runtime error", event.error || event);
  showError("Something went wrong while running the dashboard.", detail);
});

window.addEventListener("unhandledrejection", (event) => {
  const reason = event?.reason?.message || event?.reason || "unknown error";
  console.error("Unhandled promise rejection", event.reason || event);
  showError("A background task failed. Please refresh and try again.", `Details: ${reason}`);
});

let trendChartInstance = null;
let barChartInstance = null;
let extraChartInstance = null;
let heatmapChartInstance = null;
let currentTableRows = [];
let currentSort = { key: null, direction: "asc" };

const samplePath = "data-sample.csv";
const sampleManifest = [
  {
    id: "sales",
    name: "Retail Sales",
    file: "./samples/retail-sales.csv",
    description: "Daily orders, revenue, returns, and channel performance.",
    columns: ["Date", "Region", "Channel", "Orders", "Revenue", "ReturnRate"],
  },
  {
    id: "marketing",
    name: "Marketing Performance",
    file: "./samples/marketing-performance.csv",
    description: "Campaigns with spend, clicks, impressions, and conversions.",
    columns: ["Date", "Campaign", "Channel", "Spend", "Clicks", "Impressions", "Conversions", "CTR"],
  },
  {
    id: "product",
    name: "Product Usage",
    file: "./samples/product-usage.csv",
    description: "Weekly active users, sessions, and retention.",
    columns: ["Week", "Plan", "Region", "ActiveUsers", "Sessions", "RetentionRate"],
  },
];

const numberFormatter = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 2,
});

const currencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 2,
});

const percentFormatter = new Intl.NumberFormat("en-US", {
  style: "percent",
  maximumFractionDigits: 2,
});

const state = {
  rows: [],
  columns: [],
  profile: {},
  numericCandidates: [],
  dateCandidates: [],
  usableNumericMetrics: [],
  numericColumns: [],
  categoricalColumns: [],
  dateColumn: null,
  selectedMetric: null,
  selectedDimension: null,
  topN: 8,
  dateRange: { start: null, end: null },
  domain: "general",
  domainAuto: true,
  inferredDomain: "general",
  chosenXAxisType: "category",
};

document.addEventListener("DOMContentLoaded", init);

function init() {
  console.log("init started");
  if (buildStamp) {
    buildStamp.textContent = `Build: ${new Date().toISOString()}`;
  }

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });

  if (uploadTrigger && uploadInput) {
    uploadTrigger.addEventListener("click", () => uploadInput.click());
  }

  [fileInput, fileInputInline, uploadInput].forEach((input) => {
    if (!input) return;
    input.addEventListener("change", (event) => handleFileSelection(event.target.files[0]));
  });

  [loadSample, loadSampleInline].forEach((button) => {
    if (!button) return;
    button.addEventListener("click", () => loadSampleGallery());
  });

  if (dropZone) {
    dropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    });
    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
    dropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
      const file = event.dataTransfer?.files?.[0];
      handleFileSelection(file);
    });
  }

  metricSelect.addEventListener("change", () => {
    state.selectedMetric = metricSelect.value;
    updateDashboard();
  });

  dimensionSelect.addEventListener("change", () => {
    state.selectedDimension = dimensionSelect.value;
    updateDashboard();
  });

  topNSelect.addEventListener("change", () => {
    state.topN = Number(topNSelect.value);
    updateDashboard();
  });

  domainSelect.addEventListener("change", () => {
    state.domainAuto = domainSelect.value === "auto";
    state.domain = state.domainAuto ? state.inferredDomain : domainSelect.value;
    updateDashboard();
  });

  dateStartInput.addEventListener("change", () => {
    state.dateRange.start = dateStartInput.value ? new Date(dateStartInput.value) : null;
    updateDashboard();
  });

  dateEndInput.addEventListener("change", () => {
    state.dateRange.end = dateEndInput.value ? new Date(dateEndInput.value) : null;
    updateDashboard();
  });

  exportCsvButton.addEventListener("click", () => exportFilteredCsv());
  downloadInsightsButton.addEventListener("click", () => downloadInsights());

  if (testApiButton) {
    testApiButton.addEventListener("click", testApiConnection);
  }

  loadSampleGallery();
  console.log("handlers wired");
}

function handleFileSelection(file) {
  if (!file) return;
  console.log("file selected", file.name);
  clearMessages();
  setStatus("Reading file");
  const reader = new FileReader();
  reader.onload = () => {
    const text = reader.result;
    parseCsvText(text);
  };
  reader.onerror = () => {
    showError("Unable to read the file.", "Details: FileReader failed");
  };
  reader.readAsText(file);
}

function parseCsvText(text) {
  setStatus("Parsing CSV");
  Papa.parse(text, {
    header: false,
    skipEmptyLines: true,
    complete: (results) => {
      if (results.errors.length) {
        const first = results.errors[0];
        showError("We could not parse that CSV. Please check the format and try again.", `Details: ${first.message}`);
        return;
      }
      console.log("parse complete");
      ingestRows(results.data);
    },
  });
}

function loadSampleGallery() {
  renderSampleGallery();
  switchTab("samples");
}

function switchTab(tabName) {
  tabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  tabPanels.forEach((panel) => {
    panel.classList.toggle("hidden", panel.dataset.tabPanel !== tabName);
  });
}

function renderSampleGallery() {
  if (!sampleGallery) return;
  sampleGallery.innerHTML = "";
  sampleManifest.forEach((sample) => {
    const card = document.createElement("div");
    card.className = "sample-card";
    card.innerHTML = `
      <strong>${sample.name}</strong>
      <p>${sample.description}</p>
      <p class="helper-text">Columns: ${sample.columns.join(", ")}</p>
      <button class="ghost" data-sample="${sample.id}">Load sample</button>
    `;
    card.querySelector("button").addEventListener("click", () => loadSampleFile(sample));
    sampleGallery.appendChild(card);
  });
}

async function loadSampleFile(sample) {
  clearMessages();
  setStatus("Loading sample");
  try {
    const response = await fetch(sample.file);
    if (!response.ok) {
      throw new Error(`Fetch failed: ${response.status}`);
    }
    const text = await response.text();
    parseCsvText(text);
  } catch (error) {
    showError("Sample CSV could not be loaded. Check your connection or file path.", `Details: ${error.message}`);
  }
}

function testApiConnection() {
  const base = apiBaseUrl?.value?.trim();
  const endpoint = apiEndpoint?.value?.trim();
  if (!base || !endpoint) {
    apiStatus.textContent = "Please enter Base URL and Endpoint to validate.";
    return;
  }
  apiStatus.textContent = "Coming next: API connections will be supported in a future update.";
}

function clearMessages() {
  errorSection.classList.add("hidden");
  errorSection.textContent = "";
  warningsSection.classList.add("hidden");
  warningsSection.innerHTML = "";
  if (statusSection) {
    statusSection.classList.add("hidden");
    statusSection.textContent = "";
  }
}

function showError(message, detail) {
  if (detail) {
    const debug = buildDebugInfo();
    errorSection.innerHTML = `<strong>${message}</strong><div class="error-detail">${detail}</div><div class="error-detail">${debug}</div>`;
  } else {
    const debug = buildDebugInfo();
    errorSection.innerHTML = `<strong>${message}</strong><div class="error-detail">${debug}</div>`;
  }
  errorSection.classList.remove("hidden");
  console.error(message, detail || "");
}

function showWarnings(warnings) {
  if (!warnings.length) {
    warningsSection.classList.add("hidden");
    warningsSection.innerHTML = "";
    return;
  }
  warningsSection.classList.remove("hidden");
  warningsSection.innerHTML = `<strong>Data quality warnings</strong><ul>${warnings
    .map((warning) => `<li>${warning}</li>`)
    .join("")}</ul>`;
}

function showWarningMessage(message) {
  warningsSection.classList.remove("hidden");
  warningsSection.innerHTML = `<strong>Notice</strong><ul><li>${message}</li></ul>`;
}

function setStatus(message) {
  if (!statusSection) return;
  statusSection.textContent = message;
  statusSection.classList.remove("hidden");
}

function ingestRows(rawRows) {
  setStatus("Profiling columns");
  if (!rawRows || rawRows.length === 0) {
    showError("No rows found in this CSV.");
    return;
  }

  const headers = rawRows[0];
  if (!headers || headers.length === 0) {
    showError("CSV must contain a header row with column names.");
    return;
  }

  const cleanedHeaders = normalizeHeaders(headers);
  if (!cleanedHeaders.length) {
    showError("CSV must contain a header row with column names.");
    return;
  }

  const dataRows = rawRows.slice(1).filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""));
  if (!dataRows.length) {
    showError("No data rows found in this CSV.");
    return;
  }

  const rows = dataRows.map((row) => {
    const obj = {};
    cleanedHeaders.forEach((header, index) => {
      obj[header] = row[index] ?? "";
    });
    return obj;
  });

  state.rows = rows;
  state.columns = cleanedHeaders;
  state.profile = profileDataset(rows, cleanedHeaders);
  state.inferredDomain = inferDomain(state.profile);
  state.domain = state.domainAuto ? state.inferredDomain : state.domain;
  state.numericCandidates = Object.keys(state.profile).filter(
    (col) => state.profile[col].type === "numeric"
  );
  state.dateCandidates = Object.keys(state.profile).filter((col) => state.profile[col].type === "date");
  state.categoricalColumns = Object.keys(state.profile).filter(
    (col) => state.profile[col].type === "categorical"
  );
  state.usableNumericMetrics = state.numericCandidates.filter((col) => {
    let countParsed = 0;
    for (const row of rows) {
      if (parseNumber(row[col]) !== null) {
        countParsed += 1;
        if (countParsed >= 1) return true;
      }
    }
    return false;
  });
  state.numericColumns = state.usableNumericMetrics;
  state.dateColumn = state.dateCandidates[0] || null;

  if (!state.usableNumericMetrics.length) {
    const detail = buildNumericFailureDetail(state.profile);
    showError("No usable numeric metrics detected. Add at least one numeric column.", detail);
    return;
  }

  const recommendedMetrics = chooseKpiMetrics(state.profile, state.numericColumns);
  state.selectedMetric = recommendedMetrics[0] || state.numericColumns[0];
  state.selectedDimension = chooseBestDimension(state.profile, state.categoricalColumns) || state.dateColumn;

  initControls(recommendedMetrics);
  setStatus("Rendering dashboard");
  updateDashboard();

  datasetSummary.textContent = `${rows.length} rows · ${state.columns.length} columns`;
  stateSection.classList.add("hidden");
  dashboard.classList.remove("hidden");
  if (statusSection) {
    statusSection.classList.add("hidden");
    statusSection.textContent = "";
  }
  console.log("dashboard rendered");
}

function normalizeHeaders(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    let cleaned = String(header ?? "").trim();
    if (!cleaned) cleaned = `column_${index + 1}`;
    cleaned = cleaned.replace(/\s+/g, " ");
    const key = cleaned.toLowerCase();
    const count = seen.get(key) || 0;
    seen.set(key, count + 1);
    if (count > 0) {
      cleaned = `${cleaned}_${count + 1}`;
    }
    return cleaned;
  });
}

function parseNumber(value) {
  if (value === null || value === undefined) return null;
  let cleaned = String(value).trim();
  if (!cleaned) return null;

  const lower = cleaned.toLowerCase();
  if (["n/a", "na", "-", "--", ""].includes(lower)) return null;

  let isNegative = false;
  if (/^\(.*\)$/.test(cleaned)) {
    isNegative = true;
    cleaned = cleaned.replace(/^\(|\)$/g, "");
  }

  const isPercent = /%/.test(cleaned);
  cleaned = cleaned.replace(/%/g, "");

  cleaned = cleaned.replace(/\b(usd|cad|eur|gbp)\b/gi, "");
  cleaned = cleaned.replace(/[$€£]/g, "");

  cleaned = cleaned.replace(/[\s,_]/g, "");

  cleaned = cleaned.trim();
  if (!cleaned) return null;

  const number = Number(cleaned);
  if (Number.isNaN(number)) return null;
  let result = isNegative ? -number : number;
  if (isPercent) result /= 100;
  return result;
}

function parseDate(value) {
  if (!value) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const dateLike = /^\d{4}[-/]\d{2}[-/]\d{2}(?:[T\s]\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?)?$/.test(raw);
  if (!dateLike) return null;
  const parsed = Date.parse(raw);
  if (Number.isNaN(parsed)) return null;
  const date = new Date(parsed);
  if (Number.isNaN(date.valueOf())) return null;
  return date;
}

function profileDataset(rows, columns) {
  const profile = {};
  columns.forEach((col) => {
    const values = rows.map((row) => row[col]);
    const nonMissing = values.filter((value) => value !== null && value !== undefined && String(value).trim() !== "");
    const missingRate = values.length ? (values.length - nonMissing.length) / values.length : 1;

    const numericValues = [];
    const failedSamples = [];
    let numericCount = 0;
    let dateCount = 0;
    nonMissing.forEach((value) => {
      const num = parseNumber(value);
      if (num !== null) {
        numericCount += 1;
        numericValues.push(num);
      } else if (failedSamples.length < 3) {
        failedSamples.push(String(value));
      }
      const date = parseDate(value);
      if (date) dateCount += 1;
    });

    const numericRate = nonMissing.length ? numericCount / nonMissing.length : 0;
    const dateRate = nonMissing.length ? dateCount / nonMissing.length : 0;
    const nonBlankCount = nonMissing.length;

    let type = "categorical";
    if (dateRate >= 0.8) type = "date";
    else if (numericRate >= 0.6) type = "numeric";
    else if (nonBlankCount >= 20 && numericRate >= 0.7) type = "numeric";

    const uniqueSet = new Set(nonMissing.map((value) => String(value).trim()));
    const uniqueCount = uniqueSet.size;

    const stats = {
      type,
      missingRate,
      uniqueCount,
      nonNumericRate: type === "numeric" ? 1 - numericRate : null,
      numericRate,
      nonBlankCount,
      failedSamples,
      min: null,
      max: null,
      mean: null,
      median: null,
      std: null,
      dateMin: null,
      dateMax: null,
      topValues: [],
    };

    if (type === "numeric" && numericValues.length) {
      numericValues.sort((a, b) => a - b);
      stats.min = numericValues[0];
      stats.max = numericValues[numericValues.length - 1];
      const mean = numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length;
      stats.mean = mean;
      const medianIndex = Math.floor(numericValues.length / 2);
      stats.median = numericValues.length % 2 === 0
        ? (numericValues[medianIndex - 1] + numericValues[medianIndex]) / 2
        : numericValues[medianIndex];
      const variance = numericValues.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / numericValues.length;
      stats.std = Math.sqrt(variance);
    }

    if (type === "date") {
      const dates = nonMissing
        .map((value) => parseDate(value))
        .filter((value) => value !== null)
        .sort((a, b) => a - b);
      if (dates.length) {
        stats.dateMin = dates[0];
        stats.dateMax = dates[dates.length - 1];
      }
    }

    if (type === "categorical") {
      const counts = new Map();
      nonMissing.forEach((value) => {
        const key = String(value).trim() || "Unknown";
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      stats.topValues = Array.from(counts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([value, count]) => ({ value, count }));
    }

    profile[col] = stats;
  });

  return profile;
}

function inferDomain(profile) {
  const keywordSets = {
    finance: ["revenue", "sales", "cost", "spend", "profit", "margin", "expense", "cash", "balance", "income"],
    marketing: ["impression", "click", "ctr", "cvr", "roas", "campaign", "ad", "reach"],
    product: ["user", "session", "active", "retention", "engagement", "signup", "conversion"],
    operations: ["order", "shipment", "delivery", "inventory", "return", "fulfillment", "units"],
  };

  const scores = Object.keys(keywordSets).reduce((acc, key) => {
    acc[key] = 0;
    return acc;
  }, {});

  Object.keys(profile).forEach((col) => {
    const lower = col.toLowerCase();
    Object.entries(keywordSets).forEach(([domain, keywords]) => {
      if (keywords.some((keyword) => lower.includes(keyword))) {
        scores[domain] += 1;
      }
    });
  });

  const best = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
  if (!best || best[1] === 0) return "general";
  return best[0];
}

function buildNumericFailureDetail(profile) {
  const candidates = Object.entries(profile)
    .map(([col, stats]) => ({
      col,
      numericRate: stats.numericRate ?? 0,
      nonBlankCount: stats.nonBlankCount ?? 0,
      failedSamples: stats.failedSamples ?? [],
    }))
    .sort((a, b) => b.numericRate - a.numericRate)
    .slice(0, 5);

  const lines = candidates.map((item) => {
    const samples = item.failedSamples.length
      ? item.failedSamples.map((value) => escapeHtml(value)).join(", ")
      : "—";
    return `${escapeHtml(item.col)}: numericRate ${(item.numericRate * 100).toFixed(0)}%, nonBlank ${item.nonBlankCount}, samples ${samples}`;
  });

  return lines.join("<br>");
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function buildDebugInfo() {
  const numericCandidates = state.numericCandidates?.join(", ") || "—";
  const usableNumeric = state.usableNumericMetrics?.join(", ") || "—";
  const dateCandidates = state.dateCandidates?.join(", ") || "—";
  const chosenXAxisType = state.chosenXAxisType || "—";
  return `Debug: numericCandidates [${numericCandidates}], usableNumericMetrics [${usableNumeric}], dateCandidates [${dateCandidates}], chosenXAxisType ${chosenXAxisType}`;
}
function chooseKpiMetrics(profile, numericColumns) {
  const domainKeywords = {
    finance: ["revenue", "sales", "profit", "margin", "cost", "expense", "spend", "cash", "income"],
    marketing: ["impression", "click", "ctr", "cvr", "roas", "conversion", "campaign", "ad"],
    product: ["user", "session", "active", "retention", "engagement", "signup", "conversion"],
    operations: ["order", "shipment", "delivery", "inventory", "return", "units", "fulfillment"],
    general: ["revenue", "sales", "cost", "profit", "orders", "users", "sessions", "rate", "return"],
  };

  const keywords = Array.from(new Set([...(domainKeywords[state.domain] || []), ...domainKeywords.general]));

  return numericColumns
    .filter((col) => !/\b(id|guid|uuid|hash|zip|postal|code)\b/i.test(col))
    .map((col) => {
      const missingScore = 1 - (profile[col].missingRate || 0);
      const nameScore = keywords.some((keyword) => col.toLowerCase().includes(keyword)) ? 1 : 0;
      const varianceScore = profile[col].std ? Math.min(profile[col].std / (Math.abs(profile[col].mean) + 1), 1) : 0.2;
      const score = missingScore * 0.5 + nameScore * 0.35 + varianceScore * 0.15;
      return { col, score };
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, 6)
    .map((item) => item.col);
}

function chooseBestDimension(profile, categoricalColumns) {
  const filtered = categoricalColumns
    .filter((col) => profile[col].uniqueCount > 1)
    .filter((col) => profile[col].uniqueCount <= 20)
    .sort((a, b) => profile[a].uniqueCount - profile[b].uniqueCount);
  return filtered[0] || null;
}

function chooseBestCategoryForAxis(profile, categoricalColumns) {
  const filtered = categoricalColumns
    .filter((col) => profile[col].uniqueCount > 1)
    .filter((col) => profile[col].uniqueCount <= 50)
    .sort((a, b) => profile[b].uniqueCount - profile[a].uniqueCount);
  return filtered[0] || null;
}

function initControls(recommendedMetrics) {
  metricSelect.innerHTML = "";
  const allMetrics = Array.from(new Set([...recommendedMetrics, ...state.numericColumns]));
  allMetrics.forEach((metric) => {
    const option = document.createElement("option");
    option.value = metric;
    option.textContent = metric;
    metricSelect.appendChild(option);
  });
  metricSelect.value = state.selectedMetric;

  dimensionSelect.innerHTML = "";
  const dimensions = state.categoricalColumns.length ? state.categoricalColumns : [];
  if (state.dateColumn) dimensions.push(state.dateColumn);
  const uniqueDimensions = Array.from(new Set(dimensions));
  uniqueDimensions.forEach((dim) => {
    const option = document.createElement("option");
    option.value = dim;
    option.textContent = dim;
    dimensionSelect.appendChild(option);
  });
  state.selectedDimension = state.selectedDimension || uniqueDimensions[0] || state.dateColumn;
  dimensionSelect.value = state.selectedDimension;

  if (state.dateColumn) {
    dateStartInput.disabled = false;
    dateEndInput.disabled = false;
    const dateMin = state.profile[state.dateColumn].dateMin;
    const dateMax = state.profile[state.dateColumn].dateMax;
    if (dateMin && dateMax) {
      dateStartInput.min = formatDateInput(dateMin);
      dateStartInput.max = formatDateInput(dateMax);
      dateEndInput.min = formatDateInput(dateMin);
      dateEndInput.max = formatDateInput(dateMax);
      dateStartInput.value = dateStartInput.min;
      dateEndInput.value = dateEndInput.max;
      state.dateRange.start = new Date(dateStartInput.value);
      state.dateRange.end = new Date(dateEndInput.value);
    }
  } else {
    dateStartInput.disabled = true;
    dateEndInput.disabled = true;
    dateStartInput.value = "";
    dateEndInput.value = "";
    state.dateRange = { start: null, end: null };
  }

  domainSelect.value = state.domainAuto ? "auto" : state.domain;
}

function updateDashboard() {
  const filteredRows = getFilteredRows();
  renderKPIs(filteredRows);
  renderCharts(filteredRows);
  renderHeatmap(filteredRows);
  renderTable(filteredRows, state.columns);
  renderInsights(filteredRows);
  renderProfileTable(state.profile);
  renderQualityBadge(filteredRows);
  showWarnings(collectWarnings());
  renderSuggestedTrends();
}

function getFilteredRows() {
  if (!state.dateColumn) return state.rows;
  const start = state.dateRange.start;
  const end = state.dateRange.end;
  if (!start && !end) return state.rows;
  return state.rows.filter((row) => {
    const dateValue = parseDate(row[state.dateColumn]);
    if (!dateValue) return false;
    if (start && dateValue < start) return false;
    if (end && dateValue > end) return false;
    return true;
  });
}

function renderKPIs(rows) {
  kpiGrid.innerHTML = "";
  const chosenMetrics = chooseKpiMetrics(state.profile, state.numericColumns);
  const metrics = chosenMetrics.length ? chosenMetrics : state.numericColumns.slice(0, 6);

  metrics.forEach((metric) => {
    const values = rows
      .map((row) => parseNumber(row[metric]))
      .filter((value) => value !== null);
    const total = values.reduce((sum, value) => sum + value, 0);
    const avg = values.length ? total / values.length : 0;
    const median = computeMedian(values);
    const min = values.length ? Math.min(...values) : 0;
    const max = values.length ? Math.max(...values) : 0;

    const metricType = inferMetricType(metric, values);
    const primary = metricType.isRate
      ? computeRateValue(metric, values)
      : total;

    const change = computePeriodChange(rows, metric);

    const card = document.createElement("div");
    card.className = "kpi-card";
    card.innerHTML = `
      <h4>${metric}</h4>
      <div class="kpi-value">${formatMetricValue(primary, metricType)}</div>
      <div class="kpi-sub">Avg: ${formatMetricValue(avg, metricType, true)}</div>
      <div class="kpi-meta">
        <div>Median: ${formatMetricValue(median, metricType, true)}</div>
        <div>Min/Max: ${formatMetricValue(min, metricType, true)} · ${formatMetricValue(max, metricType, true)}</div>
        ${change ? `<div>${change.label}: ${change.value}</div>` : ""}
      </div>
    `;
    kpiGrid.appendChild(card);
  });
}

function renderCharts(rows) {
  if (trendChartInstance) trendChartInstance.destroy();
  if (barChartInstance) barChartInstance.destroy();
  if (extraChartInstance) extraChartInstance.destroy();

  const metric = state.selectedMetric;
  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);

  if (state.dateColumn) {
    state.chosenXAxisType = "date";
    const bucket = state.rows.length > 200 ? "week" : "day";
    const series = aggregateByDate(rows, state.dateColumn, metric, bucket, state.dateRange);
    trendTitle.textContent = `${metric} trend`;
    trendSubtitle.textContent = `Grouped by ${bucket}`;
    trendChartInstance = createLineChart("trendChart", series.labels, series.values, metricType);
  } else {
    const category = chooseBestCategoryForAxis(state.profile, state.categoricalColumns);
    if (category) {
      state.chosenXAxisType = "category";
      const series = aggregateByCategory(rows, category, metric, state.topN);
      trendTitle.textContent = `${metric} by ${category}`;
      trendSubtitle.textContent = "No date column detected";
      trendChartInstance = createLineChart("trendChart", series.labels, series.values, metricType);
    } else {
      state.chosenXAxisType = "index";
      const labels = rows.map((_, index) => index + 1);
      const values = rows.map((row) => parseNumber(row[metric])).map((value) => value ?? 0);
      trendTitle.textContent = `${metric} by row`;
      trendSubtitle.textContent = "No date column detected";
      trendChartInstance = createLineChart("trendChart", labels, values, metricType);
    }
  }

  const dimension = state.selectedDimension || chooseBestDimension(state.profile, state.categoricalColumns) || state.dateColumn;
  state.selectedDimension = dimension;
  if (dimensionSelect.value !== dimension) dimensionSelect.value = dimension;
  const breakdown = aggregateByCategory(rows, dimension, metric, state.topN);
  barTitle.textContent = `${metric} by ${dimension}`;
  barSubtitle.textContent = `Top ${state.topN}`;
  barChartInstance = createBarChart("barChart", breakdown.labels, breakdown.values, metricType);

  const correlationPair = findStrongestCorrelation(rows, state.numericColumns);
  if (correlationPair) {
    extraTitle.textContent = `Correlation: ${correlationPair.x} vs ${correlationPair.y}`;
    extraSubtitle.textContent = `r = ${correlationPair.corr.toFixed(2)}`;
    extraChartInstance = createScatterChart("extraChart", correlationPair.points);
  } else {
    extraTitle.textContent = `Distribution of ${metric}`;
    extraSubtitle.textContent = "Histogram";
    extraChartInstance = createHistogram("extraChart", metricValues);
  }
}

function renderHeatmap(rows) {
  if (heatmapChartInstance) heatmapChartInstance.destroy();
  const matrixData = buildCorrelationMatrix(rows, state.numericColumns);
  if (!matrixData || matrixData.metrics.length < 2) {
    heatmapChartInstance = createHeatmap("heatmapChart", [], []);
    return;
  }
  heatmapChartInstance = createHeatmap("heatmapChart", matrixData.metrics, matrixData.data);
}

function renderProfileTable(profile) {
  profileTable.innerHTML = "";
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["Column", "Type", "Missing %", "Unique", "Min", "Max", "Mean", "Median", "Std", "Date Range", "Top Values"]
    .forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      headerRow.appendChild(th);
    });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  Object.entries(profile).forEach(([col, stats]) => {
    const tr = document.createElement("tr");
    const dateRange = stats.dateMin && stats.dateMax
      ? `${formatDateInput(stats.dateMin)} → ${formatDateInput(stats.dateMax)}`
      : "—";
    const topValues = stats.topValues?.length
      ? stats.topValues.map((item) => `${item.value} (${item.count})`).join(", ")
      : "—";

    const metricType = stats.type === "numeric" ? inferMetricType(col, []) : { isCurrency: false, isRate: false };
    const cells = [
      col,
      stats.type,
      `${Math.round(stats.missingRate * 100)}%`,
      stats.uniqueCount ?? "—",
      stats.min !== null ? formatMetricValue(stats.min, metricType) : "—",
      stats.max !== null ? formatMetricValue(stats.max, metricType) : "—",
      stats.mean !== null ? formatMetricValue(stats.mean, metricType) : "—",
      stats.median !== null ? formatMetricValue(stats.median, metricType) : "—",
      stats.std !== null ? numberFormatter.format(stats.std) : "—",
      stats.type === "date" ? dateRange : "—",
      stats.type === "categorical" ? topValues : "—",
    ];

    cells.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  profileTable.appendChild(thead);
  profileTable.appendChild(tbody);
}

function renderQualityBadge(rows) {
  const score = computeQualityScore(rows);
  const label = `Quality ${score}`;
  qualityBadge.textContent = label;
  if (score >= 85) {
    qualityBadge.style.background = "#e8f5ef";
    qualityBadge.style.color = "#2f5e4e";
    qualityBadge.style.borderColor = "#cfe6d9";
  } else if (score >= 70) {
    qualityBadge.style.background = "#fff7e8";
    qualityBadge.style.color = "#8b5b1a";
    qualityBadge.style.borderColor = "#f3d3a2";
  } else {
    qualityBadge.style.background = "#ffe9e8";
    qualityBadge.style.color = "#a6342a";
    qualityBadge.style.borderColor = "#f7c8c2";
  }
}

function computeQualityScore(rows) {
  let score = 100;
  const profile = state.profile;
  const missingPenalty = Object.values(profile).reduce((sum, stats) => sum + stats.missingRate, 0) / state.columns.length;
  score -= Math.round(missingPenalty * 60);

  const numericIssues = Object.values(profile)
    .filter((stats) => stats.type === "numeric" && stats.nonNumericRate !== null)
    .map((stats) => stats.nonNumericRate);
  if (numericIssues.length) {
    const avgNonNumeric = numericIssues.reduce((sum, val) => sum + val, 0) / numericIssues.length;
    score -= Math.round(avgNonNumeric * 30);
  }

  const duplicateCount = countDuplicateRows(rows);
  if (duplicateCount > 0) score -= Math.min(10, Math.round((duplicateCount / rows.length) * 100));

  if (state.dateColumn) {
    const gaps = countDateGaps(rows, state.dateColumn);
    if (gaps > 0) score -= Math.min(10, Math.round((gaps / (rows.length || 1)) * 100));
  }

  return Math.max(0, Math.min(100, score));
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
    const aNum = parseNumber(aVal);
    const bNum = parseNumber(bVal);
    if (aNum !== null && bNum !== null) {
      return direction === "asc" ? aNum - bNum : bNum - aNum;
    }
    return direction === "asc"
      ? String(aVal).localeCompare(String(bVal))
      : String(bVal).localeCompare(String(aVal));
  });

  currentTableRows = sorted;
  renderTable(currentTableRows, Object.keys(currentTableRows[0] || {}));
}

function inferMetricType(metric, values) {
  const lower = metric.toLowerCase();
  const isCurrency = /(revenue|sales|cost|spend|profit|price|amount)/i.test(lower);
  const isRateName = /(rate|percent|ctr|cvr|roas|ratio)/i.test(lower);
  const isRateRange = values.length > 0 && values.every((val) => val >= 0 && val <= 1);
  const isRate = isRateName || isRateRange;
  return { isCurrency, isRate };
}

function computeRateValue(metric, values) {
  if (!values.length) return 0;
  const denominator = findRateDenominator(metric, state.numericColumns);
  if (!denominator) {
    const avg = values.reduce((sum, val) => sum + val, 0) / values.length;
    return avg;
  }
  const numerator = values.reduce((sum, val) => sum + val, 0);
  const denomValues = state.rows.map((row) => parseNumber(row[denominator])).filter((value) => value !== null);
  const denomSum = denomValues.reduce((sum, val) => sum + val, 0);
  return denomSum ? numerator / denomSum : 0;
}

function findRateDenominator(metric, numericColumns) {
  const lower = metric.toLowerCase();
  const mapping = [
    { pattern: /ctr|click.?through/, denoms: [/impression/] },
    { pattern: /cvr|conversion/, denoms: [/click/, /session/, /visit/] },
    { pattern: /return/, denoms: [/order/, /revenue/] },
    { pattern: /rate|percent|ratio/, denoms: [/order/, /click/, /impression/, /session/, /user/, /visit/] },
    { pattern: /roas/, denoms: [/spend|cost/] },
  ];

  for (const rule of mapping) {
    if (rule.pattern.test(lower)) {
      for (const denomPattern of rule.denoms) {
        const found = numericColumns.find((col) => denomPattern.test(col.toLowerCase()));
        if (found) return found;
      }
    }
  }
  return null;
}

function formatMetricValue(value, metricType) {
  if (value === null || value === undefined || Number.isNaN(value)) return "—";
  if (metricType.isRate) {
    return percentFormatter.format(value);
  }
  if (metricType.isCurrency) {
    return currencyFormatter.format(value);
  }
  return numberFormatter.format(value);
}

function computeMedian(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function computePeriodChange(rows, metric) {
  if (!state.dateColumn) return null;
  const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange);
  if (series.values.length < 6) return null;
  const window = series.values.length >= 14 ? 7 : 3;
  const recent = series.values.slice(-window);
  const previous = series.values.slice(-window * 2, -window);
  if (!previous.length) return null;
  const recentTotal = recent.reduce((sum, val) => sum + val, 0);
  const previousTotal = previous.reduce((sum, val) => sum + val, 0);
  const change = previousTotal === 0 ? 0 : (recentTotal - previousTotal) / previousTotal;
  const label = `Last ${window} vs prior ${window}`;
  const value = percentFormatter.format(change);
  return { label, value };
}

function aggregateByDate(rows, dateColumn, metric, bucket, dateRange) {
  const totals = new Map();
  rows.forEach((row) => {
    const dateValue = parseDate(row[dateColumn]);
    const metricValue = parseNumber(row[metric]);
    if (!dateValue || metricValue === null) return;
    if (dateRange.start && dateValue < dateRange.start) return;
    if (dateRange.end && dateValue > dateRange.end) return;
    const key = bucket === "week" ? getWeekStart(dateValue) : formatDateInput(dateValue);
    totals.set(key, (totals.get(key) || 0) + metricValue);
  });

  const labels = Array.from(totals.keys()).sort();
  const values = labels.map((label) => totals.get(label));
  return { labels, values };
}

function aggregateByCategory(rows, dimension, metric, topN) {
  const totals = new Map();
  rows.forEach((row) => {
    const key = row[dimension] ? String(row[dimension]).trim() : "Unknown";
    const metricValue = parseNumber(row[metric]);
    if (metricValue === null) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
  });

  const sorted = Array.from(totals.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN);

  return {
    labels: sorted.map(([label]) => label),
    values: sorted.map(([, value]) => value),
  };
}
function createLineChart(canvasId, labels, values, metricType) {
  const ctx = document.getElementById(canvasId);
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
        tooltip: {
          callbacks: {
            label: (context) => formatMetricValue(context.parsed.y, metricType),
          },
        },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: { beginAtZero: true },
      },
    },
  });
}

function createBarChart(canvasId, labels, values, metricType) {
  const ctx = document.getElementById(canvasId);
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
        tooltip: {
          callbacks: {
            label: (context) => formatMetricValue(context.parsed.y, metricType),
          },
        },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: { beginAtZero: true },
      },
    },
  });
}

function createScatterChart(canvasId, points) {
  const ctx = document.getElementById(canvasId);
  return new Chart(ctx, {
    type: "scatter",
    data: {
      datasets: [
        {
          data: points,
          backgroundColor: "rgba(47, 94, 78, 0.5)",
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true },
        y: { beginAtZero: true },
      },
    },
  });
}

function createHistogram(canvasId, values) {
  const ctx = document.getElementById(canvasId);
  const bins = 10;
  if (!values.length) {
    return new Chart(ctx, { type: "bar", data: { labels: [], datasets: [] } });
  }
  const min = Math.min(...values);
  const max = Math.max(...values);
  const step = (max - min) / bins || 1;
  const counts = Array.from({ length: bins }, () => 0);
  values.forEach((value) => {
    const index = Math.min(Math.floor((value - min) / step), bins - 1);
    counts[index] += 1;
  });
  const labels = counts.map((_, i) => `${(min + i * step).toFixed(1)}-${(min + (i + 1) * step).toFixed(1)}`);
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          data: counts,
          backgroundColor: "#f2c14e",
        },
      ],
    },
    options: { plugins: { legend: { display: false } } },
  });
}

function buildCorrelationMatrix(rows, numericColumns) {
  if (numericColumns.length < 2) return null;
  const metrics = numericColumns.filter((col) => state.profile[col].missingRate < 0.5).slice(0, 8);
  if (metrics.length < 2) return null;
  const data = [];
  metrics.forEach((xMetric, xIndex) => {
    metrics.forEach((yMetric, yIndex) => {
      const corr = computeCorrelation(rows, xMetric, yMetric);
      data.push({ x: xIndex, y: yIndex, v: corr });
    });
  });
  return { metrics, data };
}

function createHeatmap(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  return new Chart(ctx, {
    type: "matrix",
    data: {
      datasets: [
        {
          label: "Correlation",
          data,
          backgroundColor: (ctx) => {
            const value = ctx.dataset.data[ctx.dataIndex]?.v ?? 0;
            const alpha = Math.abs(value);
            return value >= 0 ? `rgba(47, 94, 78, ${alpha})` : `rgba(214, 81, 81, ${alpha})`;
          },
          borderColor: "rgba(255,255,255,0.6)",
          borderWidth: 1,
          width: (ctx) => {
            const chart = ctx.chart;
            const area = chart.chartArea;
            if (!area) return 20;
            return (area.right - area.left) / labels.length - 2;
          },
          height: (ctx) => {
            const chart = ctx.chart;
            const area = chart.chartArea;
            if (!area) return 20;
            return (area.bottom - area.top) / labels.length - 2;
          },
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: {
          type: "category",
          labels,
          offset: true,
          grid: { display: false },
        },
        y: {
          type: "category",
          labels,
          offset: true,
          grid: { display: false },
        },
      },
    },
  });
}

function findStrongestCorrelation(rows, numericColumns) {
  if (numericColumns.length < 2) return null;
  let best = null;
  for (let i = 0; i < numericColumns.length; i += 1) {
    for (let j = i + 1; j < numericColumns.length; j += 1) {
      const x = numericColumns[i];
      const y = numericColumns[j];
      const corr = computeCorrelation(rows, x, y);
      if (corr === null) continue;
      const score = Math.abs(corr);
      if (!best || score > Math.abs(best.corr)) {
        best = { x, y, corr };
      }
    }
  }
  if (!best) return null;
  const points = rows
    .map((row) => {
      const xVal = parseNumber(row[best.x]);
      const yVal = parseNumber(row[best.y]);
      if (xVal === null || yVal === null) return null;
      return { x: xVal, y: yVal };
    })
    .filter(Boolean)
    .slice(0, 500);
  return { ...best, points };
}

function computeCorrelation(rows, xMetric, yMetric) {
  const pairs = rows
    .map((row) => {
      const xVal = parseNumber(row[xMetric]);
      const yVal = parseNumber(row[yMetric]);
      if (xVal === null || yVal === null) return null;
      return [xVal, yVal];
    })
    .filter(Boolean);
  if (pairs.length < 3) return null;
  const xs = pairs.map((pair) => pair[0]);
  const ys = pairs.map((pair) => pair[1]);
  const meanX = xs.reduce((sum, val) => sum + val, 0) / xs.length;
  const meanY = ys.reduce((sum, val) => sum + val, 0) / ys.length;
  let num = 0;
  let denomX = 0;
  let denomY = 0;
  for (let i = 0; i < xs.length; i += 1) {
    const dx = xs[i] - meanX;
    const dy = ys[i] - meanY;
    num += dx * dy;
    denomX += dx * dx;
    denomY += dy * dy;
  }
  const denom = Math.sqrt(denomX * denomY);
  return denom === 0 ? null : num / denom;
}

function collectWarnings() {
  const warnings = [];
  const metric = state.selectedMetric;
  if (metric) {
    const missingRate = state.profile[metric]?.missingRate || 0;
    if (missingRate > 0.3) {
      warnings.push(`Selected metric "${metric}" has ${Math.round(missingRate * 100)}% missing values.`);
    }
    const nonNumericRate = state.profile[metric]?.nonNumericRate || 0;
    if (nonNumericRate > 0.05) {
      warnings.push(`Metric "${metric}" has ${Math.round(nonNumericRate * 100)}% non-numeric values.`);
    }
  }

  if (state.selectedDimension) {
    const uniqueCount = state.profile[state.selectedDimension]?.uniqueCount || 0;
    if (uniqueCount > 20) {
      warnings.push(`"${state.selectedDimension}" has high cardinality (${uniqueCount} unique values).`);
    }
  }

  const duplicateCount = countDuplicateRows(state.rows);
  if (duplicateCount > 0) {
    warnings.push(`${duplicateCount} duplicate rows detected.`);
  }

  if (state.dateColumn) {
    const gaps = countDateGaps(state.rows, state.dateColumn);
    if (gaps > 0) {
      warnings.push(`Detected ${gaps} missing date buckets in the trend.`);
    }
  } else {
    warnings.push("No date column detected. Charts use categories or row order.");
  }

  return warnings;
}

function countDuplicateRows(rows) {
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((row) => {
    const key = JSON.stringify(row);
    if (seen.has(key)) duplicates += 1;
    else seen.add(key);
  });
  return duplicates;
}

function countDateGaps(rows, dateColumn) {
  const dates = rows
    .map((row) => parseDate(row[dateColumn]))
    .filter(Boolean)
    .sort((a, b) => a - b);
  if (dates.length < 2) return 0;
  const days = new Set(dates.map((date) => formatDateInput(date)));
  const min = dates[0];
  const max = dates[dates.length - 1];
  let gaps = 0;
  for (let d = new Date(min); d <= max; d.setDate(d.getDate() + 1)) {
    const key = formatDateInput(d);
    if (!days.has(key)) gaps += 1;
  }
  return gaps;
}
function renderInsights(rows) {
  insightsList.innerHTML = "";
  const insights = buildInsights(rows);
  insights.slice(0, 6).forEach((text) => {
    const li = document.createElement("li");
    li.textContent = text;
    insightsList.appendChild(li);
  });
}

function renderSuggestedTrends() {
  if (!suggestedTrends) return;
  suggestedTrends.innerHTML = "";
  if (!state.rows.length) {
    suggestedTrends.innerHTML = "<li>Load a sample to see suggested trends.</li>";
    return;
  }
  const suggestions = [];
  if (state.dateColumn && state.numericColumns.length) {
    const metric = state.selectedMetric || state.numericColumns[0];
    suggestions.push(`${metric} trend over ${state.dateColumn}`);
    if (state.categoricalColumns.length) {
      suggestions.push(`${metric} by ${state.categoricalColumns[0]}`);
    }
  } else if (state.numericColumns.length && state.categoricalColumns.length) {
    suggestions.push(`${state.selectedMetric || state.numericColumns[0]} by ${state.categoricalColumns[0]}`);
  }

  if (!suggestions.length) {
    suggestions.push("Add a date or categorical column to unlock trend suggestions.");
  }

  suggestions.forEach((text) => {
    const li = document.createElement("li");
    li.textContent = text;
    suggestedTrends.appendChild(li);
  });
}

function buildInsights(rows) {
  const insights = [];
  const metric = state.selectedMetric;
  const dimension = state.selectedDimension;

  const correlationPair = findStrongestCorrelation(rows, state.numericColumns);
  if (correlationPair) {
    insights.push(
      `Strongest correlation is between ${correlationPair.x} and ${correlationPair.y} (r = ${correlationPair.corr.toFixed(2)}).`
    );
  }

  if (dimension && metric) {
    const breakdown = aggregateByCategory(rows, dimension, metric, 5);
    if (breakdown.labels.length) {
      const total = breakdown.values.reduce((sum, val) => sum + val, 0);
      const topValue = breakdown.values[0];
      const share = total ? topValue / total : 0;
      insights.push(
        `Top ${dimension} is ${breakdown.labels[0]} contributing ${percentFormatter.format(share)} of ${metric}.`
      );
      insights.push(
        `Best vs worst ${dimension}: ${breakdown.labels[0]} (${numberFormatter.format(breakdown.values[0])}) and ${breakdown.labels[breakdown.labels.length - 1]} (${numberFormatter.format(breakdown.values[breakdown.values.length - 1])}).`
      );
    }
  }

  if (state.dateColumn && metric) {
    const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange);
    const anomalies = detectAnomalies(series.values);
    if (anomalies.length) {
      const worst = anomalies[0];
      insights.push(`Largest day-over-day change in ${metric} occurs on ${series.labels[worst.index]} (z = ${worst.z.toFixed(2)}).`);
    }
  }

  Object.entries(state.profile).forEach(([col, stats]) => {
    if (stats.missingRate > 0.3) {
      insights.push(`"${col}" has ${Math.round(stats.missingRate * 100)}% missing values and may skew analysis.`);
    }
  });

  if (state.dateColumn && metric) {
    const change = computePeriodChange(rows, metric);
    if (change) {
      insights.push(`${metric} change: ${change.label} is ${change.value}.`);
    }
  }

  if (!insights.length) {
    insights.push("Upload more data to generate insights.");
  }

  return insights;
}

function detectAnomalies(values) {
  if (values.length < 4) return [];
  const changes = values.slice(1).map((val, idx) => val - values[idx]);
  const mean = changes.reduce((sum, val) => sum + val, 0) / changes.length;
  const variance = changes.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / changes.length;
  const std = Math.sqrt(variance);
  if (std === 0) return [];
  return changes
    .map((val, idx) => ({ index: idx + 1, z: (val - mean) / std }))
    .filter((item) => Math.abs(item.z) > 2)
    .sort((a, b) => Math.abs(b.z) - Math.abs(a.z));
}

function exportFilteredCsv() {
  const rows = getFilteredRows();
  if (!rows.length) return;
  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "filtered-data.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function downloadInsights() {
  const rows = getFilteredRows();
  const insights = buildInsights(rows);
  const content = [
    "Insights",
    `Generated: ${new Date().toLocaleString()}`,
    `Metric: ${state.selectedMetric}`,
    `Dimension: ${state.selectedDimension || "—"}`,
    "",
    ...insights.map((text, index) => `${index + 1}. ${text}`),
  ].join("\n");

  const blob = new Blob([content], { type: "text/plain;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "insights.txt";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getWeekStart(date) {
  const copy = new Date(date);
  const day = copy.getDay() || 7;
  copy.setDate(copy.getDate() - (day - 1));
  return formatDateInput(copy);
}
