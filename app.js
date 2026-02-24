const fileInput = document.getElementById("fileInput");
const fileInputInline = document.getElementById("fileInputInline");
const loadSample = document.getElementById("loadSample");
const loadSampleInline = document.getElementById("loadSampleInline");
const landingView = document.getElementById("landingView");
const appView = document.getElementById("appView");
const signInButton = document.getElementById("signInBtn");
const signOutButton = document.getElementById("signOutButton");
const startFreeButton = document.getElementById("startFreeBtn");
const heroStartButton = document.getElementById("heroStartBtn");
const viewSamplesButton = document.getElementById("viewSamplesBtn");
const signupModal = document.getElementById("signupModal");
const uploaderModal = document.getElementById("uploaderModal");
const closeUploaderModalButton = document.getElementById("closeUploaderModal");
const closeModalButton = document.getElementById("closeModal");
const signupForm = document.getElementById("signupForm");
const formError = document.getElementById("formError");
const leadNameInput = document.getElementById("signupName");
const leadEmailInput = document.getElementById("signupEmail");
const leadCompanyInput = document.getElementById("signupCompany");
const leadUseCaseInput = document.getElementById("signupUseCase");
const leadUseCaseDetailInput = document.getElementById("signupUseCaseDetail");
const useCaseDetailField = document.getElementById("useCaseDetailField");
const nameError = document.getElementById("nameError");
const emailError = document.getElementById("emailError");
const companyError = document.getElementById("companyError");
const useCaseError = document.getElementById("useCaseError");
const useCaseDetailError = document.getElementById("useCaseDetailError");
const signupDebug = document.getElementById("signupDebug");
const signupSubmitButton = document.getElementById("signupSubmit");
const sheetsUrlInput = document.getElementById("sheetsUrl");
const sheetsRangeInput = document.getElementById("sheetsRange");
const sheetsHelper = document.getElementById("sheetsHelper");
const loadSheetButton = document.getElementById("loadSheet");
const loginGoogleButton = document.getElementById("loginGoogle");
const googleStatus = document.getElementById("googleStatus");
const pdfInput = document.getElementById("pdfInput");
const pdfStatus = document.getElementById("pdfStatus");
const pdfMeta = document.getElementById("pdfMeta");
const exportSheetsButton = document.getElementById("exportSheets");
const downloadPdfButton = document.getElementById("downloadPdf");
const exportStatus = document.getElementById("exportStatus");
const webhookStatus = document.getElementById("webhookStatus");
const stateSection = document.getElementById("state");
const errorSection = document.getElementById("error");
const warningsSection = document.getElementById("warnings");
const statusSection = document.getElementById("status");
const metricNotice = document.getElementById("metricNotice");
const dashboard = document.getElementById("dashboard");
const dashboardSectionTitle = document.getElementById("dashboardSectionTitle");
const datasetSummary = document.getElementById("datasetSummary");
const datasetInlineNotice = document.getElementById("datasetInlineNotice");
const kpiGrid = document.getElementById("kpiGrid");
const table = document.getElementById("dataTable");
const metricSelect = document.getElementById("metricSelect");
const dimensionSelect = document.getElementById("dimensionSelect");
const topNSelect = document.getElementById("topN");
const domainSelect = document.getElementById("domainSelect");
const industrySelect = document.getElementById("industrySelect");
const industryHelper = document.getElementById("industryHelper");
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
const comparisonNote = document.getElementById("comparisonNote");
const filterBadge = document.getElementById("filterBadge");
const primaryMetricSelect = document.getElementById("primaryMetricSelect");
const compareMetricSelect = document.getElementById("compareMetricSelect");
const comparisonTrendChart = document.getElementById("comparisonTrendChart");
const driversList = document.getElementById("driversList");
const segmentChartCanvas = document.getElementById("segmentChart");
const chartsSection = document.querySelector(".charts");
const tabButtons = document.querySelectorAll(".tab-button");
const tabPanels = document.querySelectorAll("[data-tab-panel]");
const dropZone = document.getElementById("dropZone");
const uploadTrigger = document.getElementById("uploadTrigger");
const uploadInput = document.getElementById("uploadInput");
const loadFixtureButton = document.getElementById("loadFixture");
const uploadCsvQuick = document.getElementById("uploadCsvQuick");
const uploadPdfQuick = document.getElementById("uploadPdfQuick");
const uploadSheetsQuick = document.getElementById("uploadSheetsQuick");
const uploadApiQuick = document.getElementById("uploadApiQuick");
const modalDropZone = document.getElementById("modalDropZone");
const modalUploadTrigger = document.getElementById("modalUploadTrigger");
const modalCsvQuick = document.getElementById("modalCsvQuick");
const modalPdfQuick = document.getElementById("modalPdfQuick");
const modalSheetsQuick = document.getElementById("modalSheetsQuick");
const modalApiQuick = document.getElementById("modalApiQuick");
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
let comparisonChartInstance = null;
let segmentChartInstance = null;
let heatmapChartInstance = null;
let currentTableRows = [];
let currentSort = { key: null, direction: "asc" };
let googleAccessToken = null;
let googleTokenClient = null;
let logoWarningShown = false;
let renderKpiCallCount = 0;
let signupSubmitting = false;
let signupMode = "signup";
let currentUserProfile = null;
const signupFormState = {
  fullName: "",
  email: "",
  company: "",
  useCase: "",
  useCaseDetail: "",
};
const USERS_KEY = "shay_users";
const SESSION_KEY = "shay_session";

const samplePath = "data-sample.csv";
// Use the deployed Apps Script Web App URL ending in /exec (not the editor URL).
const WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbznq6aHhsHA6uiQg_AoJ6PG-WioCqriL_Z82SutiX1VeoI1TstpdqYvNPfahI8ZhwjsEQ/exec";
const GOOGLE_CLIENT_ID = "611908111462-d5mfvb991hgbfvinop8hgeec3li7rqbp.apps.googleusercontent.com";
const sampleManifest = [
  {
    id: "sales",
    name: "Retail Sales",
    file: "./samples/retail-sales.csv",
    description: "Daily orders, revenue, returns, and channel performance by industry.",
    columns: ["Date", "Industry", "Region", "Channel", "Orders", "Revenue", "ReturnRate"],
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
  normalizedDataset: null,
  datasetCapabilities: {
    hasDateField: false,
    hasNumericMetrics: false,
    hasCategoricalDimensions: false,
  },
  mode: "demo",
  rawRows: [],
  filteredRows: [],
  schema: {
    columns: [],
    profiles: {},
    numeric: [],
    dates: [],
    categoricals: [],
  },
  numericColumns: [],
  syntheticMetric: false,
  debug: {
    numericCandidates: [],
    usableNumericMetrics: [],
    dateCandidates: [],
  },
  filters: {
    industry: "All",
  },
  selections: {
    primaryMetric: null,
    compareMetrics: [],
  },
  selectedMetric: null,
  dateRange: { start: null, end: null },
  domain: "general",
  domainAuto: true,
  inferredDomain: "general",
  chosenXAxisType: "category",
  industryColumn: null,
  topN: 8,
  sampleTabAutoloaded: false,
  uiMode: "empty", // empty | loading | ready | error
};

document.addEventListener("DOMContentLoaded", init);

function init() {
  console.log("init started");
  if (buildStamp) {
    buildStamp.textContent = `Build: ${new Date().toISOString()}`;
  }

  applyBrandAssets();
  ensureUnifiedUserDashboardLayout();

  window.addEventListener("hashchange", handleRoute);

  hydrateSession();

  const isDev = location.hostname === "localhost" || location.hostname === "127.0.0.1";
  if (loadFixtureButton) {
    loadFixtureButton.classList.toggle("hidden", !isDev);
    loadFixtureButton.addEventListener("click", () => loadFixtureCsv());
  }

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });

  if (uploadTrigger && uploadInput) {
    uploadTrigger.addEventListener("click", () => uploadInput.click());
  }
  if (modalUploadTrigger && uploadInput) {
    modalUploadTrigger.addEventListener("click", () => uploadInput.click());
  }
  if (uploadCsvQuick && uploadInput) {
    uploadCsvQuick.addEventListener("click", () => uploadInput.click());
  }
  if (modalCsvQuick && uploadInput) {
    modalCsvQuick.addEventListener("click", () => uploadInput.click());
  }
  if (uploadPdfQuick && pdfInput) {
    uploadPdfQuick.addEventListener("click", () => pdfInput.click());
  }
  if (modalPdfQuick && pdfInput) {
    modalPdfQuick.addEventListener("click", () => pdfInput.click());
  }
  if (uploadSheetsQuick) {
    uploadSheetsQuick.addEventListener("click", () => switchTab("sheets"));
  }
  if (modalSheetsQuick) {
    modalSheetsQuick.addEventListener("click", () => {
      closeUploaderModal();
      switchTab("sheets");
    });
  }
  if (uploadApiQuick) {
    uploadApiQuick.addEventListener("click", () => switchTab("api"));
  }
  if (modalApiQuick) {
    modalApiQuick.addEventListener("click", () => {
      closeUploaderModal();
      switchTab("api");
    });
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
  if (modalDropZone) {
    modalDropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      modalDropZone.classList.add("dragover");
    });
    modalDropZone.addEventListener("dragleave", () => modalDropZone.classList.remove("dragover"));
    modalDropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      modalDropZone.classList.remove("dragover");
      const file = event.dataTransfer?.files?.[0];
      handleFileSelection(file);
    });
  }
  if (closeUploaderModalButton) closeUploaderModalButton.addEventListener("click", closeUploaderModal);
  if (uploaderModal) {
    uploaderModal.addEventListener("click", (event) => {
      if (event.target === uploaderModal) closeUploaderModal();
    });
  }

  metricSelect.addEventListener("change", () => {
    state.selections.primaryMetric = metricSelect.value;
    applyFiltersAndRender();
  });

  dimensionSelect.addEventListener("change", () => {
    state.selectedDimension = dimensionSelect.value;
    applyFiltersAndRender();
  });

  topNSelect.addEventListener("change", () => {
    state.topN = Number(topNSelect.value);
    applyFiltersAndRender();
  });

  domainSelect.addEventListener("change", () => {
    state.domainAuto = domainSelect.value === "auto";
    state.domain = state.domainAuto ? state.inferredDomain : domainSelect.value;
    applyFiltersAndRender();
  });

  if (industrySelect) {
    industrySelect.addEventListener("change", () => {
      state.filters.industry = industrySelect.value;
      applyFiltersAndRender();
    });
  }

  dateStartInput.addEventListener("change", () => {
    state.dateRange.start = dateStartInput.value ? new Date(dateStartInput.value) : null;
    applyFiltersAndRender();
  });

  dateEndInput.addEventListener("change", () => {
    state.dateRange.end = dateEndInput.value ? new Date(dateEndInput.value) : null;
    applyFiltersAndRender();
  });

  exportCsvButton.addEventListener("click", () => exportFilteredCsv());
  downloadInsightsButton.addEventListener("click", () => downloadInsights());

  if (testApiButton) {
    testApiButton.addEventListener("click", handleApiFetch);
  }

  if (startFreeButton) startFreeButton.addEventListener("click", handleStartFreeClick);
  if (heroStartButton) heroStartButton.addEventListener("click", handleStartFreeClick);
  if (viewSamplesButton) viewSamplesButton.addEventListener("click", () => {
    const target = document.getElementById("samples");
    if (target) target.scrollIntoView({ behavior: "smooth" });
  });
  if (closeModalButton) closeModalButton.addEventListener("click", closeModal);
  if (signupModal) signupModal.addEventListener("click", (event) => {
    if (event.target === signupModal) closeModal();
  });
  if (signupForm) signupForm.addEventListener("submit", handleSignupSubmit);
  if (leadUseCaseInput) {
    leadUseCaseInput.addEventListener("change", handleUseCaseToggle);
  }
  [leadNameInput, leadEmailInput, leadCompanyInput, leadUseCaseInput, leadUseCaseDetailInput].forEach((input) => {
    if (!input) return;
    input.addEventListener("input", handleSignupInputChange);
  });
  if (signInButton) signInButton.addEventListener("click", handleSignInClick);
  if (signOutButton) signOutButton.addEventListener("click", handleSignOutClick);
  if (loadSheetButton) loadSheetButton.addEventListener("click", loadSheetData);
  if (loginGoogleButton) loginGoogleButton.addEventListener("click", handleLoginGoogleClick);
  if (exportSheetsButton) exportSheetsButton.addEventListener("click", handleExportToSheets);
  if (downloadPdfButton) downloadPdfButton.addEventListener("click", handleExportPdf);
  if (pdfInput) {
    pdfInput.addEventListener("change", () => {
      const file = pdfInput.files?.[0];
      if (!file) return;
      if (pdfMeta) {
        const sizeKb = Math.round(file.size / 1024);
        pdfMeta.textContent = `${file.name} · ${sizeKb} KB`;
      }
      handlePdfUpload(file);
    });
  }

  updateSignInButton();
  renderSampleGallery();
  handleRoute();
  updateGoogleStatus(false);
  checkWebhookHealth();

  if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID.includes("PASTE_")) {
    console.warn("Missing GOOGLE_CLIENT_ID. Google Sheets connect will fail until you set it.");
    if (googleStatus) googleStatus.textContent = "Google connect not configured. Set GOOGLE_CLIENT_ID in app.js";
    if (loginGoogleButton) loginGoogleButton.disabled = true;
  }
  disableUnavailableButtons();
  auditActionHandlers();
  setUiMode("empty");
  updateAnalysisHeaderState(false);
  syncUploadAnalysisState();
  console.log("handlers wired");
}

function ensureUnifiedUserDashboardLayout() {
  if (!dashboard) return;
  if (dashboard.querySelector("#unifiedDashboardShell")) return;

  const charts = dashboard.querySelector(".charts");
  const insightsCard = dashboard.querySelector(".insights-card");
  const comparisonSectionEl = document.getElementById("comparisonSection");
  const tableCards = Array.from(dashboard.querySelectorAll(".table-card"));
  const controlsPanel = dashboard.querySelector(".controls-panel");
  const sectionHeader = dashboard.querySelector(".section-header");
  const demoBanner = dashboard.querySelector(".demo-banner");
  const metricNoticeEl = document.getElementById("metricNotice");
  const kpiSection = document.getElementById("kpiGrid");
  if (!charts || !insightsCard || !kpiSection) return;

  const chartCards = Array.from(charts.querySelectorAll(".chart-card"));
  const trendCard = chartCards[0];
  const performanceCard = chartCards[1];
  const extraChartCard = chartCards[2] || null;
  const filterBadgeEl = document.getElementById("filterBadge");
  if (!trendCard || !performanceCard) return;

  const shell = document.createElement("div");
  shell.id = "unifiedDashboardShell";
  shell.className = "unified-dashboard-shell";

  const row1 = document.createElement("div");
  row1.className = "unified-dashboard-row unified-dashboard-row-top";
  const kpiWrap = document.createElement("div");
  kpiWrap.className = "unified-col unified-kpi-col";
  const trendWrap = document.createElement("div");
  trendWrap.className = "unified-col unified-trend-col";

  const kpiCardWrap = document.createElement("div");
  kpiCardWrap.className = "unified-panel unified-kpi-panel";
  kpiCardWrap.appendChild(kpiSection);

  const trendPanelWrap = document.createElement("div");
  trendPanelWrap.className = "unified-panel unified-trend-panel";
  if (filterBadgeEl) trendPanelWrap.appendChild(filterBadgeEl);
  trendPanelWrap.appendChild(trendCard);

  kpiWrap.appendChild(kpiCardWrap);
  trendWrap.appendChild(trendPanelWrap);
  row1.appendChild(kpiWrap);
  row1.appendChild(trendWrap);

  const row2 = document.createElement("div");
  row2.className = "unified-dashboard-row unified-dashboard-row-middle";
  const insightsWrap = document.createElement("div");
  insightsWrap.className = "unified-col unified-insights-col";
  const perfWrap = document.createElement("div");
  perfWrap.className = "unified-col unified-performance-col";
  insightsWrap.appendChild(insightsCard);
  perfWrap.appendChild(performanceCard);
  row2.appendChild(insightsWrap);
  row2.appendChild(perfWrap);

  const evidenceRow = document.createElement("div");
  evidenceRow.id = "userEvidencePanel";
  evidenceRow.className = "unified-evidence-panel hidden";
  evidenceRow.innerHTML = `<div class="helper-text">Select an insight to view evidence details.</div>`;

  const extras = document.createElement("div");
  extras.className = "unified-dashboard-extras";
  if (extraChartCard) extras.appendChild(extraChartCard);
  if (comparisonSectionEl) extras.appendChild(comparisonSectionEl);
  tableCards.forEach((card) => extras.appendChild(card));

  shell.appendChild(row1);
  shell.appendChild(row2);
  shell.appendChild(evidenceRow);
  if (extras.children.length) shell.appendChild(extras);

  if (charts && !charts.classList.contains("hidden")) {
    charts.classList.add("hidden");
  } else {
    charts.classList.add("hidden");
  }

  const insertionTarget = controlsPanel || metricNoticeEl || sectionHeader || demoBanner;
  if (insertionTarget && insertionTarget.parentNode === dashboard) {
    insertionTarget.insertAdjacentElement("afterend", shell);
  } else {
    dashboard.appendChild(shell);
  }

  ensureDashboardActionsToolbar();
}

function ensureDashboardActionsToolbar() {
  if (!dashboard) return;
  const sectionHeader = dashboard.querySelector(".section-header");
  if (!sectionHeader) return;

  let toolbar = document.getElementById("dashboardActionsToolbar");
  if (!toolbar) {
    toolbar = document.createElement("div");
    toolbar.id = "dashboardActionsToolbar";
    toolbar.className = "dashboard-actions-toolbar";
    toolbar.innerHTML = `
      <div class="dashboard-actions-row" role="toolbar" aria-label="Dashboard actions">
        <button type="button" id="dashboardChangeDataSourceBtn" class="ghost dashboard-action-btn">Upload data</button>
      </div>
      <div id="dashboardActionsStatus" class="helper-text hidden"></div>
    `;
    sectionHeader.appendChild(toolbar);
  }

  const row = toolbar.querySelector(".dashboard-actions-row");
  const status = document.getElementById("dashboardActionsStatus");
  const changeDataButton = document.getElementById("dashboardChangeDataSourceBtn");
  if (!row) return;

  if (changeDataButton && !changeDataButton.dataset.bound) {
    changeDataButton.addEventListener("click", () => {
      openUploaderModal();
    });
    changeDataButton.dataset.bound = "true";
  }

  [downloadPdfButton, exportCsvButton, downloadInsightsButton, exportSheetsButton].forEach((button) => {
    if (!button) return;
    button.classList.add("dashboard-action-btn");
    row.appendChild(button);
  });

  if (status && exportStatus && status !== exportStatus) {
    status.replaceWith(exportStatus);
    exportStatus.id = "dashboardActionsStatus";
    exportStatus.classList.remove("hidden");
  }
  updateAnalysisHeaderState(Boolean(state.rawRows?.length));
}

function updateAnalysisHeaderState(hasDataset) {
  if (dashboardSectionTitle) {
    dashboardSectionTitle.textContent = hasDataset ? "Analysis" : "Upload";
  }
  const toolbar = document.getElementById("dashboardActionsToolbar");
  if (toolbar) toolbar.classList.remove("hidden");
  const uploadToggle = document.getElementById("dashboardChangeDataSourceBtn");
  if (uploadToggle) {
    uploadToggle.textContent = hasDataset ? "Change data source" : "Upload data";
    uploadToggle.classList.remove("hidden");
  }
  [downloadPdfButton, exportCsvButton, downloadInsightsButton, exportSheetsButton].forEach((button) => {
    if (!button) return;
    button.classList.toggle("hidden", !hasDataset);
  });
  if (!hasDataset) {
    if (datasetSummary) datasetSummary.textContent = "";
    if (datasetInlineNotice) {
      datasetInlineNotice.classList.add("hidden");
      datasetInlineNotice.textContent = "";
    }
  }
}

function openUploaderModal() {
  if (!uploaderModal) return;
  uploaderModal.classList.remove("hidden");
  uploaderModal.setAttribute("aria-hidden", "false");
}

function closeUploaderModal() {
  if (!uploaderModal) return;
  uploaderModal.classList.add("hidden");
  uploaderModal.setAttribute("aria-hidden", "true");
}

function hasLoadedDataset() {
  return Boolean(state.rawRows && state.rawRows.length > 0);
}

function hasUserDataLoaded() {
  const sourceType = state.normalizedDataset?.meta?.sourceType || null;
  if (!hasLoadedDataset()) return false;
  return sourceType !== "demo";
}

function setUiMode(mode) {
  state.uiMode = mode;
  if (mode === "ready") closeUploaderModal();
  syncUploadAnalysisState();
}

function syncUploadAnalysisState() {
  const hasDataset = hasUserDataLoaded() && state.uiMode === "ready";
  updateAnalysisHeaderState(hasDataset);

  if (dashboard) {
    dashboard.classList.remove("hidden");
    dashboard.classList.toggle("dashboard-empty-mode", !hasDataset);
  }

  const uploadTabButton = Array.from(tabButtons || []).find((button) => button.dataset.tab === "upload");
  if (uploadTabButton) {
    uploadTabButton.classList.toggle("hidden", hasDataset);
  }

  if (hasDataset) {
    tabPanels.forEach((panel) => panel.classList.add("hidden"));
  }

  const uploadTabActive = Array.from(tabButtons || []).some((button) => button.dataset.tab === "upload" && button.classList.contains("active"));
  if (stateSection) {
    stateSection.classList.add("hidden");
    stateSection.classList.toggle("is-loading", state.uiMode === "loading" && uploadTabActive);
  }

  [uploadTrigger, loadFixtureButton].forEach((btn) => {
    if (!btn) return;
    btn.disabled = state.uiMode === "loading";
  });
  if (!hasDataset) {
    if (metricNotice) metricNotice.classList.add("hidden");
    if (warningsSection) warningsSection.classList.add("hidden");
    if (statusSection && state.uiMode !== "loading") statusSection.classList.add("hidden");
  }

}

function loadUsers() {
  return JSON.parse(localStorage.getItem(USERS_KEY) || "{}");
}

function saveUsers(users) {
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
}

function setSession(email) {
  localStorage.setItem(SESSION_KEY, JSON.stringify({ email, signedInAt: new Date().toISOString() }));
}

function clearSession() {
  localStorage.removeItem(SESSION_KEY);
}

function getSession() {
  try {
    return JSON.parse(localStorage.getItem(SESSION_KEY) || "null");
  } catch (error) {
    return null;
  }
}

function hydrateSession() {
  const session = getSession();
  if (!session?.email) return;
  const users = loadUsers();
  const user = users[session.email.toLowerCase()];
  if (user) {
    currentUserProfile = user;
    signupFormState.fullName = user.fullName || "";
    signupFormState.email = user.email || session.email;
    signupFormState.company = user.company || "";
    signupFormState.useCase = user.useCase || "";
    signupFormState.useCaseDetail = user.useCaseDetail || "";
  }
}

function loadFixtureCsv() {
  const fixturePath = "./samples/fixture-stack.csv";
  clearMessages();
  setUiMode("loading");
  setStatus("Loading fixture");
  fetch(fixturePath)
    .then((response) => {
      if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
      return response.text();
    })
    .then((text) => parseCsvText(text, { sourceType: "csv", name: "fixture-stack.csv", isSample: true }))
    .catch((error) => {
      showError("Fixture CSV could not be loaded.", `Details: ${error.message}`);
    });
}

function applyBrandAssets() {
  const path = window.location.pathname || "/";
  const base = path.endsWith("/") ? path : path.replace(/\/[^/]*$/, "/");
  const logoSrc = `${base}Brand/logo.png`;
  const logo192 = `${base}Brand/logo-192.png`;
  const logo512 = `${base}Brand/logo-512.png`;

  document.querySelectorAll(".logo-img").forEach((img) => {
    img.src = logoSrc;
    img.alt = "Shay Analytics AI";
    img.onerror = () => {
      if (!logoWarningShown) {
        console.warn(`Logo failed to load: ${logoSrc}`);
        logoWarningShown = true;
      }
      img.onerror = null;
      img.src = "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='120' height='120'><rect width='100%' height='100%' fill='%2318181b'/><text x='50%' y='55%' font-size='28' text-anchor='middle' fill='%23a78bfa' font-family='Arial'>Shay</text></svg>";
    };
  });

  document.querySelectorAll("link[rel='icon']").forEach((link) => {
    link.href = link.getAttribute("sizes") === "512x512" ? logo512 : logo192;
  });
  const apple = document.querySelector("link[rel='apple-touch-icon']");
  if (apple) apple.href = logo192;
  const ogImage = document.querySelector("meta[property='og:image']");
  if (ogImage) ogImage.setAttribute("content", logo512);
  const twitterImage = document.querySelector("meta[name='twitter:image']");
  if (twitterImage) twitterImage.setAttribute("content", logo512);
}

function handleFileSelection(file) {
  if (!file) return;
  console.log("file selected", file.name);
  clearMessages();
  setUiMode("loading");
  setStatus("Reading file");
  const reader = new FileReader();
  reader.onload = () => {
    const text = reader.result;
    parseCsvText(text, {
      sourceType: "csv",
      name: file.name,
    });
  };
  reader.onerror = () => {
    showError("Unable to read the file.", "Details: FileReader failed");
  };
  reader.readAsText(file);
}

function parseCsvText(text, sourceMeta = {}) {
  setUiMode("loading");
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
      ingestRows(results.data, sourceMeta);
    },
  });
}

function loadSampleGallery() {
  renderSampleGallery();
  switchTab("samples");
}

function switchTab(tabName) {
  const isSourceTab = ["upload", "sheets", "api", "pdf"].includes(tabName);
  if (isSourceTab && hasUserDataLoaded() && state.uiMode === "ready") {
    resetStateForNewDataset();
    tabName = "upload";
  }
  if (isSourceTab && !hasUserDataLoaded()) {
    tabName = "upload";
  }
  tabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  tabPanels.forEach((panel) => {
    panel.classList.toggle("hidden", panel.dataset.tabPanel !== tabName);
  });
  syncUploadAnalysisState();
  focusSourceSection(tabName);
  if (tabName === "samples" && !state.rawRows.length && !state.sampleTabAutoloaded) {
    state.sampleTabAutoloaded = true;
    loadFixtureCsv();
  }
}

function focusSourceSection(tabName) {
  if (hasLoadedDataset() && state.uiMode === "ready") return;
  const targetMap = {
    upload: document.getElementById("upload-data"),
    sheets: document.getElementById("sheetsPanel"),
    api: document.getElementById("apiPanel"),
    pdf: document.getElementById("pdfPanel"),
    samples: document.getElementById("samplesPanel"),
  };
  const target = targetMap[tabName];
  if (!target) return;
  try {
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (_) {
    target.scrollIntoView();
  }
  if (typeof target.focus === "function") {
    target.focus({ preventScroll: true });
  }
}

function getSignedIn() {
  const session = getSession();
  if (!session?.email) return false;
  const users = loadUsers();
  if (!users[session.email.toLowerCase()]) {
    clearSession();
    return false;
  }
  return true;
}

function setSignedIn(value) {
  if (!value) {
    clearSession();
  }
}

function routeTo(viewName) {
  window.location.hash = viewName === "app" ? "#/app" : "#/landing";
}

function showView(viewName) {
  if (landingView) landingView.classList.toggle("hidden", viewName !== "landing");
  if (appView) appView.classList.toggle("hidden", viewName !== "app");
  if (viewName === "app") {
    switchTab("upload");
  }
}

function handleRoute() {
  const hash = window.location.hash || "";
  if (hash === "#/app") {
    if (!getSignedIn()) {
      routeTo("landing");
      return;
    }
    showView("app");
    return;
  }

  if (hash === "#/landing" || hash === "") {
    if (getSignedIn()) {
      routeTo("app");
      return;
    }
    showView("landing");
    return;
  }

  routeTo(getSignedIn() ? "app" : "landing");
}

function initGoogleTokenClient() {
  if (googleTokenClient) return googleTokenClient;
  if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID.includes("PASTE_")) {
    throw new Error("Missing GOOGLE_CLIENT_ID. Paste your OAuth Client ID in app.js.");
  }
  if (!window.google || !window.google.accounts || !window.google.accounts.oauth2) {
    throw new Error("Google Identity Services not available. Refresh the page.");
  }
  googleTokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly",
    callback: (tokenResponse) => {
      googleAccessToken = tokenResponse?.access_token || null;
      updateGoogleStatus(!!googleAccessToken);
      if (googleAccessToken) showToast("Google connected");
    },
  });
  return googleTokenClient;
}

function handleLoginGoogleClick() {
  try {
    const client = initGoogleTokenClient();
    client.requestAccessToken({ prompt: "consent" });
  } catch (error) {
    console.error(error);
    if (googleStatus) googleStatus.textContent = error.message || String(error);
  }
}

function updateGoogleStatus(isLoggedIn) {
  if (googleStatus) {
    googleStatus.textContent = isLoggedIn ? "Connected" : "Not connected";
    googleStatus.classList.toggle("offline", !isLoggedIn);
  }
}

function updateSignInButton() {
  if (!signInButton) return;
  signInButton.textContent = getSignedIn() ? "Dashboard" : "Sign In";
}

function handleSignInClick() {
  if (getSignedIn()) {
    routeTo("app");
    return;
  }
  signupMode = "signin";
  openModal();
}

function handleStartFreeClick() {
  signupMode = "signup";
  openModal();
}

function handleSignOutClick() {
  setSignedIn(false);
  currentUserProfile = null;
  signupMode = "signup";
  updateSignInButton();
  routeTo("landing");
}

function openModal() {
  if (!signupModal) return;
  signupModal.classList.remove("hidden");
  signupModal.setAttribute("aria-hidden", "false");
  if (formError) {
    formError.textContent = "";
    formError.classList.add("hidden");
  }
  if (nameError) nameError.textContent = "";
  if (emailError) emailError.textContent = "";
  if (companyError) companyError.textContent = "";
  if (useCaseError) useCaseError.textContent = "";
  if (useCaseDetailError) useCaseDetailError.textContent = "";
  hydrateSignupForm();
  toggleSignupFields();
  handleUseCaseToggle();
  updateSignupButtonState();
}

function closeModal() {
  if (!signupModal) return;
  signupModal.classList.add("hidden");
  signupModal.setAttribute("aria-hidden", "true");
}

function handleSignupSubmit(event) {
  event.preventDefault();
  const validation = validateSignupForm();
  if (!validation.isValid) {
    updateSignupButtonState();
    return;
  }

  signupSubmitting = true;
  updateSignupButtonState();

  if (signupMode === "signin") {
    const profile = findStoredProfile(validation.cleaned.email);
    if (!profile) {
      if (formError) formError.textContent = "No account found. Create an account.";
      signupSubmitting = false;
      updateSignupButtonState();
      return;
    }
    signupFormState.fullName = profile.fullName || "";
    signupFormState.company = profile.company || "";
    signupFormState.useCase = profile.useCase || "";
    signupFormState.useCaseDetail = profile.useCaseDetail || "";
    signupFormState.email = profile.email || validation.cleaned.email;
    currentUserProfile = profile;
    setSession(validation.cleaned.email.toLowerCase());
    setSignedIn(true);
    logEvent({
      source: "auth",
      action: "sign_in",
      fullName: profile.fullName || "",
      email: validation.cleaned.email,
      company: profile.company || "",
      useCase: profile.useCase || "",
      useCaseDetail: profile.useCaseDetail || "",
      timestamp: new Date().toISOString(),
    }).catch(() => {});
    signupSubmitting = false;
    closeModal();
    showToast("Signed in");
    routeTo("app");
    return;
  }

  const { name, email, company, useCase, useCaseDetail } = validation.cleaned;
  const normalizedEmail = email.trim().toLowerCase();

  const lead = {
    name,
    email,
    company,
    useCase,
    useCaseDetail,
    source: "shay analytics ai landing",
    userAgent: navigator.userAgent,
    createdAt: new Date().toISOString(),
  };

  const payload = {
    fullName: name,
    email: normalizedEmail,
    company,
    useCase,
    useCaseDetail: useCaseDetail || "",
    source: "signup",
    action: "sign_up",
    timestamp: new Date().toISOString(),
  };

  postSignupToAppsScript(payload).then((result) => {
    if (formError) {
      formError.textContent = result.ok
        ? "Account created. You can now access the dashboard."
        : "Account created, but logging failed. Please try again.";
    }
    signupSubmitting = false;
    updateSignupButtonState();
  });

  saveLead(lead);
  const users = loadUsers();
  users[normalizedEmail] = {
    fullName: name,
    email: normalizedEmail,
    company,
    useCase,
    useCaseDetail: useCaseDetail || "",
    createdAt: new Date().toISOString(),
  };
  saveUsers(users);
  setSession(normalizedEmail);
  currentUserProfile = users[normalizedEmail];
  setSignedIn(true);
  updateSignInButton();
  closeModal();
  showToast("Account created. You can now access the dashboard.");
  routeTo("app");
}

function hydrateSignupForm() {
  if (leadNameInput) leadNameInput.value = signupFormState.fullName;
  if (leadEmailInput) leadEmailInput.value = signupFormState.email;
  if (leadCompanyInput) leadCompanyInput.value = signupFormState.company;
  if (leadUseCaseInput) leadUseCaseInput.value = signupFormState.useCase;
  if (leadUseCaseDetailInput) leadUseCaseDetailInput.value = signupFormState.useCaseDetail;
}

function handleSignupInputChange() {
  signupFormState.fullName = leadNameInput?.value || "";
  signupFormState.email = leadEmailInput?.value || "";
  signupFormState.company = leadCompanyInput?.value || "";
  signupFormState.useCase = leadUseCaseInput?.value || "";
  signupFormState.useCaseDetail = leadUseCaseDetailInput?.value || "";
  handleUseCaseToggle();
  updateSignupButtonState();
}

function handleUseCaseToggle() {
  if (signupMode !== "signup") {
    if (useCaseDetailField) useCaseDetailField.classList.add("hidden");
    return;
  }
  const useCase = (leadUseCaseInput?.value || "").trim();
  if (useCaseDetailField) {
    const show = useCase === "Other";
    useCaseDetailField.classList.toggle("hidden", !show);
  }
  updateSignupButtonState();
}

function validateSignupForm() {
  const name = (leadNameInput?.value || "").trim();
  const email = (leadEmailInput?.value || "").trim();
  const company = (leadCompanyInput?.value || "").trim();
  const useCase = (leadUseCaseInput?.value || "").trim();
  const useCaseDetail = (leadUseCaseDetailInput?.value || "").trim();

  const errors = {};
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) errors.email = "Enter a valid email.";
  if (signupMode === "signup") {
    if (name.length < 2) errors.name = "Full name is required.";
    if (company.length < 2) errors.company = "Company is required.";
    if (!useCase) errors.useCase = "Select a use case.";
    if (useCase === "Other" && useCaseDetail.length < 10) {
      errors.useCaseDetail = "Please describe your use case (min 10 characters).";
    }
  }

  if (nameError) nameError.textContent = errors.name || "";
  if (emailError) emailError.textContent = errors.email || "";
  if (companyError) companyError.textContent = errors.company || "";
  if (useCaseError) useCaseError.textContent = errors.useCase || "";
  if (useCaseDetailError) useCaseDetailError.textContent = errors.useCaseDetail || "";

  return {
    isValid: Object.keys(errors).length === 0,
    cleaned: { name, email, company, useCase, useCaseDetail },
  };
}

function updateSignupButtonState() {
  if (!signupSubmitButton) return;
  const validation = validateSignupForm();
  signupSubmitButton.disabled = !validation.isValid || signupSubmitting;
  if (signupDebug) {
    const isDev = location.hostname === "localhost" || location.hostname === "127.0.0.1";
    signupDebug.classList.toggle("hidden", !isDev);
    if (isDev) {
      signupDebug.textContent = `isValid=${validation.isValid} isSubmitting=${signupSubmitting} | name=${validation.cleaned.name.length} email=${validation.cleaned.email.length} company=${validation.cleaned.company.length} useCase="${validation.cleaned.useCase}" detail=${validation.cleaned.useCaseDetail.length}`;
    }
  }
}

function toggleSignupFields() {
  document.querySelectorAll(".signup-only").forEach((el) => {
    el.classList.toggle("hidden-field", signupMode !== "signup");
  });
  const title = signupModal?.querySelector("h3");
  const helper = signupModal?.querySelector(".helper-text");
  if (title) title.textContent = signupMode === "signup" ? "Start for Free" : "Sign in";
  if (helper) {
    helper.textContent = signupMode === "signup"
      ? "Create your account to unlock the dashboard."
      : "Enter your email to access the dashboard.";
  }
  if (signupSubmitButton) {
    signupSubmitButton.textContent = signupMode === "signup" ? "Create account" : "Sign in";
  }
}

function findStoredProfile(email) {
  const users = loadUsers();
  return users[email.toLowerCase()];
}

function showFormError(message) {
  if (!formError) return;
  formError.textContent = message;
  formError.classList.remove("hidden");
}

function saveLead(lead) {
  const existing = JSON.parse(localStorage.getItem("leads") || "[]");
  existing.unshift(lead);
  localStorage.setItem("leads", JSON.stringify(existing));
}

function showStatus(msg) {
  const el = document.getElementById("signup-status");
  if (el) el.textContent = msg;
}

async function postSignupToAppsScript(payload) {
  if (!WEBHOOK_URL) {
    return { ok: false, status: 0, body: "missing webhook" };
  }

  const res = await fetch(WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });

  const text = await res.text();
  let ok = res.ok;
  try {
    const json = JSON.parse(text);
    ok = !!json.ok;
  } catch (e) {}

  return { ok, status: res.status, body: text };
}

function showToast(message) {
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add("show"));
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 300);
  }, 1800);
}

function disableUnavailableButtons() {
  if (exportSheetsButton) {
    exportSheetsButton.disabled = true;
    exportSheetsButton.title = "Coming soon: Google Sheets export";
    if (exportStatus) exportStatus.textContent = "Export to Google Sheets coming soon.";
  }
}

function auditActionHandlers() {
  const handlerMap = new Map([
    ["start-free", handleStartFreeClick],
    ["signin", handleSignInClick],
    ["view-samples", () => {}],
    ["tab-upload", () => {}],
    ["tab-sheets", () => {}],
    ["tab-api", () => {}],
    ["tab-pdf", () => {}],
    ["tab-samples", () => {}],
    ["signout", handleSignOutClick],
    ["export-pdf", handleExportPdf],
    ["test-api", handleApiFetch],
    ["load-sheet", loadSheetData],
    ["login-google", handleLoginGoogleClick],
    ["export-csv", exportFilteredCsv],
    ["download-insights", downloadInsights],
    ["export-sheets", handleExportToSheets],
  ]);

  document.querySelectorAll("[data-action]").forEach((button) => {
    const action = button.getAttribute("data-action");
    if (!handlerMap.has(action) && !button.disabled) {
      console.warn(`Missing handler for data-action="${action}"`);
    }
  });
}

async function logEvent(payload) {
  if (!WEBHOOK_URL) return { ok: false, status: 0, body: "missing webhook" };
  try {
    const response = await fetch(WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    const text = await response.text();
    return { ok: response.ok, status: response.status, body: text };
  } catch (error) {
    return { ok: false, status: 0, body: String(error) };
  }
}

function checkWebhookHealth() {
  if (!WEBHOOK_URL || !webhookStatus) return;
  fetch(WEBHOOK_URL)
    .then((response) => response.text().then((text) => ({ response, text })))
    .then(({ response, text }) => {
      let ok = response.ok;
      try {
        const json = JSON.parse(text);
        ok = !!json.ok;
      } catch (e) {}
      webhookStatus.textContent = ok
        ? "Webhook online"
        : "Webhook offline. Verify the Apps Script Web App URL and deployment.";
    })
    .catch(() => {
      webhookStatus.textContent = "Webhook offline. Verify the Apps Script Web App URL and deployment.";
    });
}

async function handleExportPdf() {
  const target = document.getElementById("dashboard");
  if (!target || target.classList.contains("hidden")) {
    showError("Load a dataset before exporting the report.");
    return;
  }
  if (!window.html2canvas || !window.jspdf) {
    showError("PDF export library not available. Please refresh and try again.");
    return;
  }

  showToast("Preparing PDF...");
  const canvas = await window.html2canvas(target, { scale: 2, backgroundColor: "#ffffff" });
  const imgData = canvas.toDataURL("image/png");
  const pdf = new window.jspdf.jsPDF("p", "pt", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 24;
  const imgWidth = pageWidth - margin * 2;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;

  pdf.setFontSize(14);
  pdf.text("Shay Analytics AI Report", margin, margin);
  pdf.setFontSize(10);
  pdf.text(new Date().toLocaleString(), margin, margin + 14);

  let position = margin + 24;
  let remainingHeight = imgHeight;

  while (remainingHeight > 0) {
    pdf.addImage(imgData, "PNG", margin, position, imgWidth, imgHeight);
    remainingHeight -= pageHeight - margin * 2;
    if (remainingHeight > 0) {
      pdf.addPage();
      position = margin - (imgHeight - remainingHeight) + margin * 2;
    }
  }

  const dateStamp = new Date().toISOString().slice(0, 10);
  pdf.save(`shalytics-report-${dateStamp}.pdf`);
}

async function handleApiFetch() {
  const base = apiBaseUrl?.value?.trim();
  const endpoint = apiEndpoint?.value?.trim();
  if (!base || !endpoint) {
    apiStatus.textContent = "Please enter Base URL and Endpoint to fetch data.";
    return;
  }
  const url = `${base.replace(/\/$/, "")}/${endpoint.replace(/^\//, "")}`;
  const authType = apiAuthType?.value || "none";
  const credential = apiKey?.value?.trim();
  const headers = {};
  if (authType === "bearer" && credential) {
    headers.Authorization = `Bearer ${credential}`;
  } else if (authType === "apiKey" && credential) {
    headers["x-api-key"] = credential;
  }

  clearMessages();
  setStatus("Fetching API data");
  try {
    const response = await fetch(url, { headers });
    if (!response.ok) {
      apiStatus.textContent = `Request failed: ${response.status}`;
      showError("API request failed.", `Details: ${response.status}`);
      return;
    }
    const json = await response.json();
    if (!Array.isArray(json) || !json.length || typeof json[0] !== "object") {
      apiStatus.textContent = "Expected a JSON array of objects.";
      showError("API payload must be a JSON array of objects.");
      return;
    }
    apiStatus.textContent = "API data loaded";
    ingestJsonRows(json, { sourceType: "api", name: `${base}${endpoint}` });
    switchTab("upload");
  } catch (error) {
    apiStatus.textContent = "Network error while fetching API data.";
    showError("API request failed.", `Details: ${error.message}`);
  }
}

function ingestJsonRows(rows, sourceMeta = {}) {
  const columns = Array.from(
    rows.reduce((set, row) => {
      Object.keys(row || {}).forEach((key) => set.add(key));
      return set;
    }, new Set())
  );
  const rawRows = [columns, ...rows.map((row) => columns.map((col) => row[col]))];
  ingestRows(rawRows, sourceMeta);
}

function handleExportToSheets() {
  const user = currentUserProfile?.email || "anonymous";
  const filters = {};
  if (state.filters?.industry && state.filters.industry !== "All") {
    filters.industry = state.filters.industry;
  }
  if (state.dateRange?.start) filters.dateStart = formatDateInput(state.dateRange.start);
  if (state.dateRange?.end) filters.dateEnd = formatDateInput(state.dateRange.end);
  if (state.selectedMetric) filters.metric = state.selectedMetric;
  if (state.selectedDimension) filters.dimension = state.selectedDimension;

  if (exportStatus) {
    exportStatus.textContent = "Export request sent. Confirm in shared sheet within 10 seconds.";
  }

  const payload = {
    source: "web_app",
    user,
    action: "export_google_sheet",
    entity: "dashboard",
    value: state.selectedMetric || "Dashboard",
    filters,
    app: "Shay Analytics AI",
    ts: new Date().toISOString(),
  };

  logEvent(payload).then((result) => {
    if (!result.ok) {
      const message = result.body || "Export request failed.";
      if (exportStatus) exportStatus.textContent = message;
    }
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
  setUiMode("loading");
  setStatus("Loading sample");
  try {
    const response = await fetch(sample.file);
    if (!response.ok) {
      throw new Error(`Fetch failed: ${response.status}`);
    }
    const text = await response.text();
    parseCsvText(text, { sourceType: "csv", name: sample?.name || sample?.file || "sample.csv", isSample: true });
  } catch (error) {
    showError("Sample CSV could not be loaded. Check your connection or file path.", `Details: ${error.message}`);
  }
}

function activateSourceTab(sourceMeta = {}) {
  const sourceType = String(sourceMeta?.sourceType || "").toLowerCase();
  if (sourceMeta?.isSample) {
    switchTab("samples");
    return;
  }
  if (sourceType === "sheet") {
    switchTab("sheets");
    return;
  }
  if (sourceType === "api") {
    switchTab("api");
    return;
  }
  if (sourceType === "pdf") {
    switchTab("pdf");
    return;
  }
  switchTab("upload");
}

function normalizeSheetsUrl(input) {
  const url = input?.trim() || "";
  if (!url) return "";
  if (!url.includes("docs.google.com/spreadsheets")) return url;
  if (url.includes("export?format=csv")) return url;
  const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return url;
  const gidMatch = url.match(/gid=([0-9]+)/);
  const gid = gidMatch ? gidMatch[1] : "0";
  return `https://docs.google.com/spreadsheets/d/${idMatch[1]}/export?format=csv&gid=${gid}`;
}

function extractSheetId(input) {
  const url = input?.trim() || "";
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(url)) return url;
  return "";
}

function buildSheetsFallback(url) {
  if (!url.includes("docs.google.com/spreadsheets")) return "";
  const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return "";
  const gidMatch = url.match(/gid=([0-9]+)/);
  const gid = gidMatch ? gidMatch[1] : "0";
  return `https://docs.google.com/spreadsheets/d/${idMatch[1]}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

async function fetchSheetText(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
  return response.text();
}

async function loadSheetData() {
  const input = sheetsUrlInput?.value || "";
  const range = sheetsRangeInput?.value?.trim() || "A1:Z1000";
  clearMessages();
  setUiMode("loading");
  setStatus("Loading Google Sheet");
  try {
    const spreadsheetId = extractSheetId(input);
    if (googleAccessToken && spreadsheetId) {
      const apiRange = encodeURIComponent(range || "A1:Z1000");
      const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${apiRange}?majorDimension=ROWS`;
      const response = await fetch(apiUrl, {
        headers: { Authorization: `Bearer ${googleAccessToken}` },
      });
      if (response.status === 401 || response.status === 403) {
        updateGoogleStatus(false);
        showError("Permission denied or token expired. Reconnect Google.");
        return;
      }
      if (!response.ok) {
        const body = await response.text();
        showError("Google Sheets read failed.", `Details: ${body}`);
        return;
      }
      const payload = await response.json();
      if (!payload?.values || !payload.values.length) {
        showError("No rows returned from Google Sheets.");
        return;
      }
      updateGoogleStatus(true);
      const csvText = valuesToCsv(payload.values || []);
      parseCsvText(csvText, { sourceType: "sheet", name: spreadsheetId || "Google Sheet" });
      switchTab("upload");
      return;
    }

    const normalizedUrl = normalizeSheetsUrl(input);
    if (!normalizedUrl) {
      showError("Paste a Google Sheets link first.");
      return;
    }
    const text = await fetchSheetText(normalizedUrl);
    parseCsvText(text, { sourceType: "sheet", name: normalizedUrl });
    switchTab("upload");
  } catch (error) {
    const fallback = buildSheetsFallback(normalizeSheetsUrl(input) || input);
    if (fallback) {
      try {
        const text = await fetchSheetText(fallback);
        parseCsvText(text, { sourceType: "sheet", name: fallback });
        switchTab("upload");
        return;
      } catch (fallbackError) {
        console.warn("Fallback CSV fetch failed", fallbackError);
      }
    }
    const detail = error?.message || String(error);
    if (detail.includes("Failed to fetch")) {
      showError("This sheet is likely private. Publish it or share a public CSV export link.");
    } else {
      showError("Google Sheets CSV could not be loaded. Check the link and sharing settings.", `Details: ${detail}`);
    }
  }
}

function valuesToCsv(values) {
  if (!values.length) return "";
  return values.map((row) => row.map(csvEscape).join(",")).join("\n");
}

async function handlePdfUpload(file) {
  if (!file) return;
  setUiMode("loading");
  if (pdfStatus) pdfStatus.textContent = "Extracting table data from PDF...";
  try {
    if (!window.pdfjsLib) {
      showError("PDF parser not available. Please refresh and try again.");
      return;
    }
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const maxPages = Math.min(6, pdf.numPages || 0);
    const pageAnalyses = [];

    for (let pageNum = 1; pageNum <= maxPages; pageNum += 1) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const items = (textContent.items || [])
        .filter((item) => String(item.str || "").trim() !== "")
        .map((item) => ({
          text: String(item.str || "").trim(),
          x: item.transform[4],
          y: item.transform[5],
          width: item.width || 0,
          height: item.height || 0,
        }));
      const pageRows = reconstructPdfPageRows(items);
      const pageTables = detectPdfTablesFromRows(pageRows);
      pageAnalyses.push({ pageNum, items, rows: pageRows, tables: pageTables });
    }

    const selectedTable = selectBestPdfFactTable(pageAnalyses);
    const fallbackRows = pageAnalyses.flatMap((page) => page.rows.map((row) => row.cells || []));
    const finalRows = selectedTable?.rows?.length ? selectedTable.rows : fallbackRows;

    if (finalRows.length < 2) {
      showError("No table like content detected. Export a CSV or use a PDF with a clear table.");
      if (pdfStatus) pdfStatus.textContent = "PDF ingestion failed";
      return;
    }

    const csvText = finalRows.map((row) => row.map(csvEscape).join(",")).join("\n");
    parseCsvText(csvText, {
      sourceType: "pdf",
      name: file?.name || "PDF upload",
      pdfMode: "table",
      extractedTableKind: selectedTable?.kind || "generic",
    });
    if (pdfStatus) pdfStatus.textContent = "PDF parsed. Generating dashboard...";
    switchTab("upload");
  } catch (error) {
    showError("No table like content detected. Export a CSV or use a PDF with a clear table.", `Details: ${error.message}`);
    if (pdfStatus) pdfStatus.textContent = "PDF ingestion failed";
  }
}

function reconstructPdfPageRows(items) {
  const yTolerance = 3;
  const rows = [];
  const sortedItems = [...(items || [])].sort((a, b) => {
    const yDiff = b.y - a.y;
    if (Math.abs(yDiff) > yTolerance) return yDiff;
    return a.x - b.x;
  });

  sortedItems.forEach((item) => {
    let target = rows.find((row) => Math.abs(row.y - item.y) <= yTolerance);
    if (!target) {
      target = { y: item.y, items: [] };
      rows.push(target);
    }
    target.items.push(item);
  });

  rows.sort((a, b) => b.y - a.y);
  const xAnchors = buildPdfColumnAnchors(rows);

  return rows.map((row) => {
    const cells = Array.from({ length: xAnchors.length }, () => "");
    row.items
      .sort((a, b) => a.x - b.x)
      .forEach((item) => {
        const idx = nearestAnchorIndex(item.x, xAnchors);
        if (idx < 0) return;
        cells[idx] = cells[idx] ? `${cells[idx]} ${item.text}`.trim() : item.text;
      });
    const compact = cells.map((cell) => String(cell || "").trim());
    return { y: row.y, cells: compact };
  }).filter((row) => row.cells.filter(Boolean).length >= 2);
}

function buildPdfColumnAnchors(rows) {
  const xs = [];
  (rows || []).forEach((row) => {
    (row.items || []).forEach((item) => {
      if (item.x != null) xs.push(item.x);
    });
  });
  xs.sort((a, b) => a - b);
  const anchors = [];
  const tolerance = 14;
  xs.forEach((x) => {
    const last = anchors[anchors.length - 1];
    if (!last || Math.abs(last - x) > tolerance) anchors.push(x);
  });
  return anchors.slice(0, 18);
}

function nearestAnchorIndex(x, anchors) {
  if (!anchors?.length) return -1;
  let bestIdx = 0;
  let best = Infinity;
  anchors.forEach((anchor, idx) => {
    const diff = Math.abs(anchor - x);
    if (diff < best) {
      best = diff;
      bestIdx = idx;
    }
  });
  return bestIdx;
}

function detectPdfTablesFromRows(pageRows) {
  const blocks = [];
  let current = [];
  (pageRows || []).forEach((row) => {
    const nonEmpty = row.cells.filter(Boolean).length;
    if (nonEmpty >= 3) {
      current.push(row);
    } else if (current.length) {
      blocks.push(current);
      current = [];
    }
  });
  if (current.length) blocks.push(current);

  return blocks
    .map((block) => buildPdfTableCandidate(block))
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
}

function buildPdfTableCandidate(block) {
  if (!Array.isArray(block) || block.length < 2) return null;
  const rows = block.map((row) => row.cells);
  const headerIndex = findPdfHeaderRowIndex(rows);
  if (headerIndex < 0 || headerIndex >= rows.length - 1) return null;
  let header = normalizePdfHeaderRow(rows[headerIndex]);
  let dataRows = rows.slice(headerIndex + 1).filter((row) => row.some((cell) => String(cell || "").trim() !== ""));
  if (!header.length || dataRows.length < 1) return null;

  // Trim columns that are entirely empty in data.
  const keepIdx = header.map((_, idx) => {
    if (String(header[idx] || "").trim()) return true;
    return dataRows.some((row) => String(row[idx] || "").trim());
  });
  header = header.filter((_, idx) => keepIdx[idx]);
  dataRows = dataRows.map((row) => row.filter((_, idx) => keepIdx[idx]));

  const normalized = normalizePdfTableRows(header, dataRows);
  const score = scorePdfTableCandidate(normalized.header, normalized.rows);
  if (score <= 0) return null;
  return {
    header: normalized.header,
    dataRows: normalized.rows,
    rows: [normalized.header, ...normalized.rows],
    score,
    kind: scorePdfTableKind(normalized.header),
  };
}

function findPdfHeaderRowIndex(rows) {
  let bestIdx = -1;
  let bestScore = -Infinity;
  (rows || []).slice(0, 8).forEach((row, idx) => {
    const cells = (row || []).map((cell) => String(cell || "").trim()).filter(Boolean);
    if (cells.length < 2) return;
    const marketingHits = cells.reduce((sum, cell) => sum + (isMarketingMetricToken(cell) ? 1 : 0), 0);
    const dateHits = cells.reduce((sum, cell) => sum + (isLikelyTemporalColumnName(cell) || isMonthToken(cell) ? 1 : 0), 0);
    const numericLike = cells.reduce((sum, cell) => sum + (parseNumber(cell) !== null ? 1 : 0), 0);
    const textDensity = cells.length - numericLike;
    const score = marketingHits * 5 + dateHits * 3 + textDensity * 1 - numericLike * 0.75;
    if (score > bestScore) {
      bestScore = score;
      bestIdx = idx;
    }
  });
  return bestScore >= 2 ? bestIdx : 0;
}

function normalizePdfHeaderRow(headerRow) {
  const header = (headerRow || []).map((cell, idx) => {
    const raw = String(cell || "").trim();
    if (!raw) return `Column_${idx + 1}`;
    const cleaned = raw.replace(/\s+/g, " ").replace(/[^\w\s%/-]/g, "").trim();
    return cleaned || `Column_${idx + 1}`;
  });
  const seen = new Map();
  return header.map((name) => {
    const key = name.toLowerCase();
    const count = seen.get(key) || 0;
    seen.set(key, count + 1);
    return count ? `${name}_${count + 1}` : name;
  });
}

function normalizePdfTableRows(header, rows) {
  let outHeader = [...header];
  let outRows = rows.map((row) => [...row]);

  // If months appear as headers, unpivot to long format: Metric | Month | Value
  const monthHeaderIndices = outHeader
    .map((name, idx) => ({ name, idx }))
    .filter((item) => isMonthToken(item.name));
  if (monthHeaderIndices.length >= 3 && outHeader.length >= 4) {
    const metricColIndex = outHeader.findIndex((name, idx) => idx !== monthHeaderIndices[0].idx && !monthHeaderIndices.some((m) => m.idx === idx));
    if (metricColIndex >= 0) {
      const longRows = [];
      outRows.forEach((row) => {
        const metricName = String(row[metricColIndex] || "").trim();
        if (!metricName) return;
        monthHeaderIndices.forEach((month) => {
          const value = row[month.idx];
          if (String(value || "").trim() === "") return;
          longRows.push([metricName, outHeader[month.idx], value]);
        });
      });
      if (longRows.length >= 6) {
        outHeader = ["Metric", "Month", "Value"];
        outRows = longRows;
      }
    }
  }
  return { header: outHeader, rows: outRows };
}

function scorePdfTableCandidate(header, rows) {
  const headerCells = (header || []).map((h) => String(h).toLowerCase());
  const marketingHits = headerCells.filter((cell) => isMarketingMetricToken(cell)).length;
  const dateHits = headerCells.filter((cell) => isLikelyTemporalColumnName(cell) || isMonthToken(cell)).length;
  const numericDensity = estimatePdfTableNumericDensity(rows);
  const rowCountScore = Math.min((rows?.length || 0) / 5, 3);
  return (marketingHits * 6) + (dateHits * 3) + (numericDensity * 4) + rowCountScore;
}

function scorePdfTableKind(header) {
  const headerCells = (header || []).map((h) => String(h).toLowerCase());
  const marketingHits = headerCells.filter((cell) => isMarketingMetricToken(cell)).length;
  return marketingHits >= 3 ? "campaign_fact" : "generic_table";
}

function estimatePdfTableNumericDensity(rows) {
  let total = 0;
  let numeric = 0;
  (rows || []).slice(0, 20).forEach((row) => {
    (row || []).forEach((cell) => {
      const text = String(cell || "").trim();
      if (!text) return;
      total += 1;
      if (parseNumber(text) !== null) numeric += 1;
    });
  });
  return total ? numeric / total : 0;
}

function isMarketingMetricToken(text) {
  const key = String(text || "").toLowerCase();
  return [
    "spend",
    "media cost",
    "mediacost",
    "impressions",
    "clicks",
    "ctr",
    "cpc",
    "conversions",
    "secondary conversions",
    "ecpm",
    "cpa",
    "attributed sales",
    "revenue",
  ].some((token) => key.includes(token));
}

function isMonthToken(text) {
  const raw = String(text || "").trim();
  if (!raw) return false;
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)$/i.test(raw)) return true;
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\s+\d{2,4}$/i.test(raw)) return true;
  return false;
}

function selectBestPdfFactTable(pageAnalyses) {
  const candidates = (pageAnalyses || []).flatMap((page) => (page.tables || []).map((table) => ({
    ...table,
    pageNum: page.pageNum,
  })));
  if (!candidates.length) return null;
  const primaryFact = candidates.find((candidate) => candidate.kind === "campaign_fact");
  return primaryFact || candidates[0];
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, "\"\"")}"`;
  }
  return text;
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

function resetStateForNewDataset() {
  destroyAllCharts();
  state.filteredRows = [];
  state.schema = { columns: [], profiles: {}, numeric: [], dates: [], categoricals: [] };
  state.numericColumns = [];
  state.debug = { numericCandidates: [], usableNumericMetrics: [], dateCandidates: [] };
  state.filters = { industry: "All" };
  state.selections = { primaryMetric: null, compareMetrics: [] };
  state.selectedMetric = null;
  state.industryColumn = null;
  state.dateRange = { start: null, end: null };
  state.chosenXAxisType = "category";
  state.normalizedDataset = null;
  state.uiMode = "empty";
  updateAnalysisHeaderState(false);
  syncUploadAnalysisState();
  if (kpiGrid) kpiGrid.innerHTML = "";
  if (driversList) driversList.innerHTML = "";
  if (insightsList) insightsList.innerHTML = "";
  if (profileTable) profileTable.innerHTML = "";
  if (table) table.innerHTML = "";
  if (chartsSection) {
    chartsSection.querySelectorAll("canvas").forEach((canvas) => {
      const ctx = canvas.getContext("2d");
      if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
    });
  }
  if (filterBadge) filterBadge.textContent = "Filtered to: All";
}

function destroyAllCharts() {
  [trendChartInstance, barChartInstance, extraChartInstance, comparisonChartInstance, segmentChartInstance, heatmapChartInstance]
    .filter(Boolean)
    .forEach((chart) => chart.destroy());
  trendChartInstance = null;
  barChartInstance = null;
  extraChartInstance = null;
  comparisonChartInstance = null;
  segmentChartInstance = null;
  heatmapChartInstance = null;
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
  if (!hasLoadedDataset()) {
    state.uiMode = "error";
    syncUploadAnalysisState();
  }
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

/**
 * @typedef {{
 * rows: Record<string, any>[],
 * columns: { name: string, type: "number" | "currency" | "percent" | "date" | "string" }[],
 * meta: {
 *   sourceType: "csv" | "pdf" | "sheet" | "api" | "demo",
 *   name?: string,
 *   dateField?: string,
 *   hasDateField: boolean,
 *   hasNumericMetrics: boolean,
 *   hasCategoricalDimensions: boolean,
 *   pdfMode?: "table" | "text"
 * }
 * }} NormalizedDataset
 */

function inferNormalizedColumnType(columnName, profile) {
  if (!profile) return "string";
  if (profile.type === "date") return "date";
  if (profile.type === "numeric") {
    const lower = String(columnName || "").toLowerCase();
    if (lower.includes("rate") || lower.includes("percent") || lower.includes("%") || lower.includes("ctr")) return "percent";
    if (lower.includes("revenue") || lower.includes("spend") || lower.includes("amount") || lower.includes("price") || lower.includes("cost")) {
      return "currency";
    }
    return "number";
  }
  return "string";
}

function isLikelyTemporalColumnName(columnName) {
  const key = String(columnName || "").trim().toLowerCase();
  return [
    "date",
    "time",
    "timestamp",
    "datetime",
    "month",
    "period",
    "week",
    "created_at",
    "updated_at",
    "run_date",
    "batch",
  ].some((token) => key === token || key.includes(token));
}

function canParseTemporalValue(value) {
  if (value === null || value === undefined) return false;
  const text = String(value).trim();
  if (!text) return false;
  if (parseDate(text)) return true;

  // Common month/period formats not always captured by parseDate consistently.
  if (/^\d{4}[-/]\d{1,2}$/.test(text)) return true; // YYYY-MM
  if (/^[A-Za-z]{3,9}\s+\d{4}$/.test(text)) return true; // Jan 2024 / January 2024
  if (/^\d{4}\s*W\d{1,2}$/i.test(text) || /^W\d{1,2}\s+\d{4}$/i.test(text)) return true; // 2024W12 / W12 2024
  if (/^Q[1-4]\s+\d{4}$/i.test(text) || /^\d{4}\s+Q[1-4]$/i.test(text)) return true; // Q1 2024 / 2024 Q1
  return false;
}

function inferFlexibleDateField({ rows, columns, profiles, fallbackDateField }) {
  if (fallbackDateField) return fallbackDateField;
  const candidateColumns = (columns || []).filter((col) => {
    const profile = profiles?.[col];
    if (profile?.type === "date") return true;
    return isLikelyTemporalColumnName(col);
  });

  for (const col of candidateColumns) {
    const sampleValues = (rows || [])
      .map((row) => row?.[col])
      .filter((value) => value !== null && value !== undefined && String(value).trim() !== "")
      .slice(0, 25);
    if (!sampleValues.length) continue;
    const parsedCount = sampleValues.filter((value) => canParseTemporalValue(value)).length;
    if (parsedCount >= Math.max(2, Math.ceil(sampleValues.length * 0.4))) {
      return col;
    }
  }
  return null;
}

function isLikelyIdColumnName(columnName) {
  const key = String(columnName || "").trim().toLowerCase();
  return key === "id" || key.endsWith("_id") || key.includes("uuid");
}

function buildNormalizedDataset({ rows, columns, profiles, sourceMeta = {} }) {
  const normalizedColumns = (columns || []).map((name) => ({
    name,
    type: inferNormalizedColumnType(name, profiles?.[name]),
  }));
  const inferredDateField = inferFlexibleDateField({
    rows,
    columns,
    profiles,
    fallbackDateField: sourceMeta.dateField || state.dateColumn || undefined,
  }) || undefined;
  const hasDateField = Boolean(
    inferredDateField
    || normalizedColumns.some((column) => column.type === "date" || isLikelyTemporalColumnName(column.name))
  );
  const hasNumericMetrics = normalizedColumns.some((column) => column.type === "number" || column.type === "currency" || column.type === "percent");
  const hasCategoricalDimensions = normalizedColumns.some((column) => column.type === "string" && !isLikelyIdColumnName(column.name));
  /** @type {NormalizedDataset} */
  const dataset = {
    rows: Array.isArray(rows) ? rows : [],
    columns: normalizedColumns,
    meta: {
      sourceType: sourceMeta.sourceType || "csv",
      name: sourceMeta.name || undefined,
      dateField: inferredDateField,
      hasDateField,
      hasNumericMetrics,
      hasCategoricalDimensions,
      pdfMode: sourceMeta.pdfMode || undefined,
    },
  };
  return dataset;
}

function getDashboardModeForDataset(dataset) {
  if (!dataset) return "demo";
  if (dataset.meta?.sourceType === "pdf" && dataset.meta?.pdfMode === "text") return "pdf-text";
  return dataset.meta?.sourceType === "demo" ? "demo" : "user";
}

function renderDashboardRenderer({ dataset, mode }) {
  // Shared renderer entry point. For now it delegates to the existing dashboard layout/render pipeline.
  // All sources should call into this function after normalization.
  if (!dataset) return;
  state.normalizedDataset = dataset;
  state.datasetCapabilities = {
    hasDateField: Boolean(dataset.meta?.hasDateField),
    hasNumericMetrics: Boolean(dataset.meta?.hasNumericMetrics),
    hasCategoricalDimensions: Boolean(dataset.meta?.hasCategoricalDimensions),
  };
  state.mode = mode;
  state.uiMode = dataset?.rows?.length ? "ready" : "empty";
  dashboard?.setAttribute("data-dashboard-mode", mode || "user");
  dashboard?.setAttribute("data-source-type", dataset.meta?.sourceType || "csv");
  if (dataset.meta && dataset.meta.hasDateField === false) {
    state.dateColumn = null;
  } else if (dataset.meta?.dateField) {
    state.dateColumn = dataset.meta.dateField;
  }
  ensureUnifiedUserDashboardLayout();
  ensureDashboardActionsToolbar();
  dashboard?.classList.add("dashboard-unified-active");
  updateAnalysisHeaderState(Boolean(dataset?.rows?.length));
  syncUploadAnalysisState();
  applyFiltersAndRender();
}

function ingestRows(rawRows, sourceMeta = {}) {
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

  resetStateForNewDataset();
  state.rawRows = rows;
  state.schema.columns = cleanedHeaders;
  state.schema.profiles = profileDataset(rows, cleanedHeaders);
  state.inferredDomain = inferDomain(state.schema.profiles);
  state.domain = state.domainAuto ? state.inferredDomain : state.domain;
  const numericCandidates = Object.keys(state.schema.profiles)
    .filter((col) => state.schema.profiles[col].type === "numeric")
    .filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], rows.length));
  const dateCandidates = Object.keys(state.schema.profiles).filter((col) => state.schema.profiles[col].type === "date");
  const categoricalCandidates = Object.keys(state.schema.profiles).filter(
    (col) => state.schema.profiles[col].type === "categorical"
  );
  const usableNumericMetrics = numericCandidates.filter((col) => {
    let countParsed = 0;
    for (const row of state.rawRows) {
      if (parseNumber(row[col]) !== null) {
        countParsed += 1;
        if (countParsed >= 1) return true;
      }
    }
    return false;
  });
  state.schema.numeric = Array.isArray(usableNumericMetrics) ? usableNumericMetrics : [];
  state.numericColumns = state.schema.numeric;
  state.schema.dates = Array.isArray(dateCandidates) ? dateCandidates : [];
  state.schema.categoricals = Array.isArray(categoricalCandidates) ? categoricalCandidates : [];
  state.debug.numericCandidates = numericCandidates;
  state.debug.usableNumericMetrics = state.schema.numeric;
  state.debug.dateCandidates = state.schema.dates;
  state.industryColumn = findIndustryColumn(state.schema.columns);
  state.dateColumn = chooseBestDateColumn(state.schema.profiles);

  ensureMetrics(state.schema, state.rawRows);

  const recommendedMetrics = chooseKpiMetrics(state.schema.profiles, state.schema.numeric);
  state.selections.primaryMetric = recommendedMetrics[0] || state.schema.numeric[0] || null;
  state.selectedMetric = state.selections.primaryMetric;
  state.selectedDimension = chooseBestDimension(state.schema.profiles, state.schema.categoricals) || state.schema.dates[0] || null;
  state.normalizedDataset = buildNormalizedDataset({
    rows: state.rawRows,
    columns: state.schema.columns,
    profiles: state.schema.profiles,
    sourceMeta: {
      ...sourceMeta,
      dateField: state.dateColumn,
    },
  });
  state.mode = getDashboardModeForDataset(state.normalizedDataset);

  initControls(recommendedMetrics);
  setStatus("Rendering dashboard");
  renderDashboardRenderer({
    dataset: state.normalizedDataset,
    mode: state.mode,
  });

  const sourceBadge = state.normalizedDataset?.meta?.sourceType ? ` · ${String(state.normalizedDataset.meta.sourceType).toUpperCase()}` : "";
  datasetSummary.textContent = `${rows.length} rows · ${state.schema.columns.length} columns${sourceBadge}`;
  stateSection.classList.add("hidden");
  dashboard.classList.remove("hidden");
  activateSourceTab(sourceMeta);
  syncUploadAnalysisState();
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

  if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(cleaned)) return null;
  const number = Number(cleaned);
  if (Number.isNaN(number)) return null;
  let result = isNegative ? -number : number;
  if (isPercent) result /= 100;
  return result;
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return Number.isNaN(value.valueOf()) ? null : value;
  const raw = String(value).trim();
  if (!raw) return null;

  const digitsOnly = /^[0-9]+$/.test(raw);
  if (digitsOnly) return null;

  if (raw.includes("T") || raw.endsWith("Z")) {
    const isoDate = new Date(raw);
    if (!Number.isNaN(isoDate.valueOf())) return isoDate;
  }

  const ymd = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (ymd) {
    const year = Number(ymd[1]);
    const month = Number(ymd[2]);
    const day = Number(ymd[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return date;
    }
  }

  const mdy = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) {
    const month = Number(mdy[1]);
    const day = Number(mdy[2]);
    const year = Number(mdy[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return date;
    }
  }

  const ym = raw.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (ym) {
    const year = Number(ym[1]);
    const month = Number(ym[2]);
    const date = new Date(Date.UTC(year, month - 1, 1));
    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1) {
      return date;
    }
  }

  const monthYear = raw.match(/^([A-Za-z]{3,9})\s+(\d{4})$/);
  if (monthYear) {
    const parsed = new Date(`${monthYear[1]} 1, ${monthYear[2]}`);
    if (!Number.isNaN(parsed.valueOf())) return parsed;
  }

  const isoWeekA = raw.match(/^(\d{4})\s*W(\d{1,2})$/i);
  const isoWeekB = raw.match(/^W(\d{1,2})\s+(\d{4})$/i);
  const isoWeek = isoWeekA || isoWeekB;
  if (isoWeek) {
    const year = Number(isoWeekA ? isoWeekA[1] : isoWeekB[2]);
    const week = Number(isoWeekA ? isoWeekA[2] : isoWeekB[1]);
    if (week >= 1 && week <= 53) {
      const jan4 = new Date(Date.UTC(year, 0, 4));
      const jan4Day = jan4.getUTCDay() || 7;
      const week1Monday = new Date(jan4);
      week1Monday.setUTCDate(jan4.getUTCDate() - (jan4Day - 1));
      const result = new Date(week1Monday);
      result.setUTCDate(week1Monday.getUTCDate() + ((week - 1) * 7));
      return result;
    }
  }

  const quarterA = raw.match(/^Q([1-4])\s+(\d{4})$/i);
  const quarterB = raw.match(/^(\d{4})\s+Q([1-4])$/i);
  const quarter = quarterA || quarterB;
  if (quarter) {
    const year = Number(quarterA ? quarterA[2] : quarterB[1]);
    const q = Number(quarterA ? quarterA[1] : quarterB[2]);
    return new Date(Date.UTC(year, (q - 1) * 3, 1));
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.valueOf()) ? null : fallback;
}

function profileDataset(rows, columns) {
  const profile = {};
  columns.forEach((col) => {
    const values = rows.map((row) => row[col]);
    const nonMissing = values.filter((value) => value !== null && value !== undefined && String(value).trim() !== "");
    const missingRate = values.length ? (values.length - nonMissing.length) / values.length : 1;

    const numericValues = [];
    const dateValues = [];
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
      if (date) {
        dateCount += 1;
        dateValues.push(date);
      }
    });

    const numericRate = nonMissing.length ? numericCount / nonMissing.length : 0;
    const dateRate = nonMissing.length ? dateCount / nonMissing.length : 0;
    const nonBlankCount = nonMissing.length;

    const lower = col.toLowerCase();
    let type = "categorical";
    if (dateRate >= 0.9 || (dateRate >= 0.8 && numericRate < 0.5)) type = "date";
    else if (numericRate >= 0.9) type = "numeric";

    if (/activeusers|sessions/.test(lower)) type = "numeric";

    const uniqueSet = new Set(nonMissing.map((value) => String(value).trim()));
    const uniqueCount = uniqueSet.size;
    const uniqueRate = nonBlankCount ? uniqueCount / nonBlankCount : 0;

    const confidenceScore = type === "numeric"
      ? numericRate
      : type === "date"
        ? dateRate
        : Math.max(0.3, 1 - Math.max(numericRate, dateRate));

    const stats = {
      type,
      inferredType: type,
      confidenceScore,
      parseFailures: nonBlankCount - (type === "date" ? dateCount : numericCount),
      examples: nonMissing.slice(0, 3).map((value) => String(value)),
      missingRate,
      uniqueCount,
      uniqueRate,
      nonNumericRate: type === "numeric" ? 1 - numericRate : null,
      numericRate,
      dateRate,
      nonBlankCount,
      failedSamples,
      integerOnly: null,
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
      stats.integerOnly = numericValues.every((value) => Number.isInteger(value));
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
      const dates = dateValues.sort((a, b) => a - b);
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
  const numericCandidates = (state.debug.numericCandidates || []).join(", ") || "—";
  const usableNumeric = (state.debug.usableNumericMetrics || []).join(", ") || "—";
  const dateCandidates = (state.debug.dateCandidates || []).join(", ") || "—";
  const chosenXAxisType = state.chosenXAxisType || "—";
  return `Debug: numericCandidates [${numericCandidates}], usableNumericMetrics [${usableNumeric}], dateCandidates [${dateCandidates}], chosenXAxisType ${chosenXAxisType}`;
}
function chooseKpiMetrics(profile, numericColumns) {
  return numericColumns
    .map((col) => ({ col, quality: classifyNumericMetricColumn(col, profile[col], state.rawRows?.length || 0) }))
    .filter((item) => !item.quality.isIdentifierLike)
    .map((col) => {
      const missingScore = 1 - (profile[col.col].missingRate || 0);
      const varianceScore = profile[col.col].std ? Math.min(profile[col.col].std / (Math.abs(profile[col.col].mean) + 1), 1) : 0.2;
      const typePriority = col.quality.kindPriority;
      const score = (typePriority * 0.55) + (missingScore * 0.3) + (varianceScore * 0.15);
      return { col: col.col, score };
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, 6)
    .map((item) => item.col);
}

function formatMetricLabel(metric) {
  return metric === "__row_count__" ? "Row Count" : metric;
}

function ensureMetrics(schema, rows) {
  const numeric = schema.numeric || [];
  if (numeric.length > 0) {
    state.syntheticMetric = false;
    if (metricNotice) {
      metricNotice.classList.add("hidden");
      metricNotice.textContent = "";
    }
    console.assert(numeric.length > 0, "Numeric metrics should not be empty when real metrics exist.");
    return numeric;
  }

  const syntheticId = "__row_count__";
  if (!schema.columns.includes(syntheticId)) {
    schema.columns.push(syntheticId);
  }
  rows.forEach((row) => {
    row[syntheticId] = 1;
  });
  schema.profiles[syntheticId] = {
    type: "numeric",
    missingRate: 0,
    uniqueCount: 1,
    uniqueRate: 0,
    nonNumericRate: 0,
    numericRate: 1,
    nonBlankCount: rows.length,
    failedSamples: [],
    integerOnly: true,
    min: 1,
    max: 1,
    mean: 1,
    median: 1,
    std: 0,
    dateMin: null,
    dateMax: null,
    topValues: [],
  };
  schema.numeric = [syntheticId];
  state.numericColumns = schema.numeric;
  state.syntheticMetric = true;
  if (metricNotice) {
    metricNotice.textContent = "No numeric columns detected. Using Row Count as the primary metric.";
    metricNotice.classList.remove("hidden");
  }
  console.assert(schema.numeric.length === 1 && schema.numeric[0] === syntheticId, "Synthetic metric injection failed.");
  return schema.numeric;
}

function chooseBestDimension(profile, categoricalColumns) {
  const filtered = categoricalColumns
    .filter((col) => profile[col].uniqueCount > 1)
    .filter((col) => profile[col].uniqueCount <= 20)
    .sort((a, b) => profile[a].uniqueCount - profile[b].uniqueCount);
  return filtered[0] || null;
}

function chooseCategoryColumn(rows, profiles) {
  const columns = state.schema.columns?.length ? state.schema.columns : Object.keys(profiles || {});
  const categoricals = columns.filter((col) => profiles[col]?.type === "categorical");
  const within = categoricals.filter((col) => profiles[col]?.uniqueCount >= 2 && profiles[col]?.uniqueCount <= 25);
  if (within.length) return within[0];
  const sorted = categoricals
    .filter((col) => profiles[col]?.uniqueCount > 1)
    .sort((a, b) => (profiles[a]?.uniqueCount || Infinity) - (profiles[b]?.uniqueCount || Infinity));
  return sorted[0] || null;
}

function isIdLikeColumn(colName, profile, rowCount) {
  const lower = String(colName || "").toLowerCase();
  const namePatterns = [
    "id",
    "uuid",
    "guid",
    "account id",
    "account_id",
    "user id",
    "user_id",
    "campaign id",
    "campaign_id",
    "order id",
    "order_id",
  ];
  if (namePatterns.some((pattern) => lower.includes(pattern))) return true;
  const uniqueRate = profile?.uniqueRate ?? 0;
  if (rowCount >= 20 && uniqueRate >= 0.9) {
    const looksIdNamed = namePatterns.some((pattern) => lower.includes(pattern));
    if (looksIdNamed) return true;

    if (profile?.integerOnly) {
      const nonBlank = profile.nonBlankCount || rowCount || 0;
      const range = (profile?.max !== null && profile?.min !== null) ? Math.abs(profile.max - profile.min) : 0;
      if (nonBlank >= 20 && uniqueRate >= 0.98 && range > nonBlank * 10) return true;
    }
  }
  if (profile?.integerOnly && profile?.min !== null && profile?.max !== null) {
    const range = Math.abs(profile.max - profile.min);
    const nonBlank = profile.nonBlankCount || rowCount || 0;
    if (nonBlank >= 20 && range > nonBlank * 10 && uniqueRate >= 0.8) return true;
  }
  return false;
}

function findIndustryColumn(columns) {
  const normalized = columns.map((c) => ({ raw: c, key: String(c).trim().toLowerCase() }));
  const exact = normalized.find((c) => c.key === "industry");
  if (exact) return exact.raw;

  const candidates = ["industry", "vertical", "sector", "category"];
  const contains = normalized.find((c) => candidates.some((k) => c.key.includes(k)));
  return contains ? contains.raw : null;
}

function chooseBestDateColumn(profiles) {
  const candidates = Object.entries(profiles)
    .filter(([_, p]) => p.type === "date")
    .sort((a, b) => (b[1].dateRate || 0) - (a[1].dateRate || 0));
  return candidates.length ? candidates[0][0] : null;
}

function initControls(recommendedMetrics) {
  metricSelect.innerHTML = "";
  const allMetrics = Array.from(new Set([...(recommendedMetrics || []), ...(state.schema.numeric || [])]));
  (allMetrics || []).forEach((metric) => {
    const option = document.createElement("option");
    option.value = metric;
    option.textContent = formatMetricLabel(metric);
    metricSelect.appendChild(option);
  });
  if (!state.selections.primaryMetric || !allMetrics.includes(state.selections.primaryMetric)) {
    state.selections.primaryMetric = allMetrics[0] || null;
  }
  state.selectedMetric = state.selections.primaryMetric;
  metricSelect.value = state.selections.primaryMetric || "";

  dimensionSelect.innerHTML = "";
  const dimensions = state.schema.categoricals?.length ? state.schema.categoricals : [];
  if (state.schema.dates[0]) dimensions.push(state.schema.dates[0]);
  const uniqueDimensions = Array.from(new Set(dimensions));
  uniqueDimensions.forEach((dim) => {
    const option = document.createElement("option");
    option.value = dim;
    option.textContent = dim;
    dimensionSelect.appendChild(option);
  });
  state.selectedDimension = state.selectedDimension || uniqueDimensions[0] || state.schema.dates[0] || null;
  dimensionSelect.value = state.selectedDimension;

  if (state.dateColumn) {
    dateStartInput.disabled = false;
    dateEndInput.disabled = false;
    const dateMin = state.schema.profiles[state.dateColumn]?.dateMin;
    const dateMax = state.schema.profiles[state.dateColumn]?.dateMax;
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

  if (industrySelect) {
    const previous = state.filters.industry || "All";
    industrySelect.innerHTML = "";
    const allOption = document.createElement("option");
    allOption.value = "All";
    allOption.textContent = "All";
    industrySelect.appendChild(allOption);

    if (state.industryColumn) {
      industrySelect.disabled = false;
      const seen = new Map();
      state.rawRows.forEach((row) => {
        const rawValue = row[state.industryColumn];
        if (rawValue === null || rawValue === undefined) return;
        const trimmed = String(rawValue).trim();
        if (!trimmed) return;
        const key = trimmed.toLowerCase();
        if (!seen.has(key)) seen.set(key, trimmed);
      });
      const values = Array.from(seen.values()).sort((a, b) => a.localeCompare(b));
      values.forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        industrySelect.appendChild(option);
      });
      if (values.includes(previous)) {
        industrySelect.value = previous;
        state.filters.industry = previous;
      } else {
        industrySelect.value = "All";
        state.filters.industry = "All";
      }
      if (industryHelper) industryHelper.textContent = "";
    } else {
      industrySelect.disabled = true;
      industrySelect.value = "All";
      state.filters.industry = "All";
      if (industryHelper) industryHelper.textContent = "Industry not found in this file";
    }
  }
}

function applyFiltersAndRender() {
  clearMessages();
  state.filteredRows = applyIndustryFilter(state.rawRows, state.filters.industry);
  let scopedRows = getFilteredRows();
  if ((!scopedRows || scopedRows.length === 0) && state.filters.industry !== "All") {
    state.filters.industry = "All";
    if (industrySelect) industrySelect.value = "All";
    state.filteredRows = applyIndustryFilter(state.rawRows, state.filters.industry);
    scopedRows = getFilteredRows();
  }
  if (!Array.isArray(scopedRows) || scopedRows.length === 0) {
    showError("No rows match this filter.");
    return;
  }

  state.schema.profiles = profileDataset(scopedRows, state.schema.columns);
  const numericCandidates = Object.keys(state.schema.profiles)
    .filter((col) => state.schema.profiles[col].type === "numeric")
    .filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], scopedRows.length));
  const dateCandidates = Object.keys(state.schema.profiles).filter((col) => state.schema.profiles[col].type === "date");
  const categoricalCandidates = Object.keys(state.schema.profiles).filter((col) => state.schema.profiles[col].type === "categorical");
  const usableNumeric = (numericCandidates || []).filter((col) => {
    let countParsed = 0;
    for (const row of scopedRows) {
      if (parseNumber(row[col]) !== null) {
        countParsed += 1;
        if (countParsed >= 1) return true;
      }
    }
    return false;
  });
  state.schema.numeric = Array.isArray(usableNumeric) ? usableNumeric : [];
  state.numericColumns = state.schema.numeric;
  state.schema.dates = Array.isArray(dateCandidates) ? dateCandidates : [];
  state.schema.categoricals = Array.isArray(categoricalCandidates) ? categoricalCandidates : [];
  state.debug.numericCandidates = numericCandidates;
  state.debug.usableNumericMetrics = state.schema.numeric;
  state.debug.dateCandidates = state.schema.dates;
  if (!state.dateColumn) {
    state.dateColumn = chooseBestDateColumn(state.schema.profiles);
  }
  ensureMetrics(state.schema, scopedRows);
  if (!state.selections.primaryMetric || !state.schema.numeric.includes(state.selections.primaryMetric)) {
    state.selections.primaryMetric = state.schema.numeric[0] || null;
  }
  state.selectedMetric = state.selections.primaryMetric;
  state.selections.compareMetrics = (state.selections.compareMetrics || []).filter((m) => state.schema.numeric.includes(m));

  if (topNSelect) {
    const uniqueCount = state.selectedDimension ? (state.schema.profiles[state.selectedDimension]?.uniqueCount || 0) : 0;
    if (uniqueCount && uniqueCount <= 2) {
      topNSelect.disabled = true;
      topNSelect.title = "Top N disabled for 2-category breakdowns.";
    } else {
      topNSelect.disabled = false;
      topNSelect.title = "";
    }
  }

  if (filterBadge) {
    filterBadge.textContent = `Filtered to: ${state.filters.industry || "All"}`;
  }
  if (datasetInlineNotice) {
    if (!state.schema.dates.length) {
      datasetInlineNotice.textContent = "Time trend unavailable, no date field detected";
      datasetInlineNotice.classList.remove("hidden");
    } else {
      datasetInlineNotice.classList.add("hidden");
      datasetInlineNotice.textContent = "";
    }
  }

  if (!runStep("buildKPIs", () => renderKPIs(scopedRows))) return;
  if (!runStep("buildCharts", () => renderCharts(scopedRows))) return;
  if (!runStep("buildComparison", () => renderMetricComparison(scopedRows))) return;
  if (!runStep("buildTable", () => renderTable(scopedRows, state.schema.columns))) return;
  if (!runStep("buildInsights", () => renderInsights(scopedRows))) return;
  runStep("buildProfile", () => renderProfileTable(state.schema.profiles));
  runStep("buildQuality", () => renderQualityBadge(scopedRows));
  runStep("buildWarnings", () => showWarnings(collectWarnings()));
  runStep("buildSuggestions", () => renderSuggestedTrends());
  updateIngestionStatus(scopedRows);
}

function runStep(step, fn) {
  try {
    const summary = buildSchemaSummary();
    console.log(`STEP: ${step}`, summary);
    fn();
    console.log(`STEP DONE: ${step}`);
    return true;
  } catch (error) {
    console.error(`STEP FAILED: ${step}`, error);
    showError(`Dashboard step failed: ${step}.`, `Details: ${error?.message || error}`);
    if (step === "buildKPIs") {
      showError("KPI generation failed. Open Debug panel for details.", buildNumericFailureDetail(state.schema.profiles));
    }
    return false;
  }
}

function buildSchemaSummary() {
  return {
    rowCount: state.filteredRows?.length || 0,
    selectedMetric: state.selections.primaryMetric,
    selectedDateField: state.dateColumn,
    selectedBreakdown: state.selectedDimension,
    numericCandidates: state.debug.numericCandidates,
    dateCandidates: state.debug.dateCandidates,
  };
}

function updateIngestionStatus(rows) {
  if (!statusSection) return;
  const numeric = state.schema.numeric || [];
  const dateColumn = state.dateColumn || "—";
  const excludedIds = (state.debug.numericCandidates || []).filter((col) => !numeric.includes(col));
  const rowCount = rows?.length || 0;
  const primaryMetric = formatMetricLabel(state.selections.primaryMetric || numeric[0] || "—");
  statusSection.innerHTML = `
    <strong>Ingestion status</strong><br>
    Rows: ${rowCount} · Date column: ${dateColumn} · Primary metric: ${primaryMetric}<br>
    Metrics: ${numeric.map(formatMetricLabel).slice(0, 6).join(", ") || "—"}<br>
    Excluded IDs: ${excludedIds.join(", ") || "—"}
  `;
  statusSection.classList.remove("hidden");
}

function applyIndustryFilter(rows, industry) {
  if (!state.industryColumn || !industry || industry === "All") return rows;
  const target = String(industry).trim().toLowerCase();
  return rows.filter((row) => String(row[state.industryColumn] ?? "").trim().toLowerCase() === target);
}

function getFilteredRows() {
  if (!state.dateColumn) return state.filteredRows || [];
  const start = state.dateRange.start;
  const end = state.dateRange.end;
  if (!start && !end) return state.filteredRows || [];
  return (state.filteredRows || []).filter((row) => {
    const dateValue = parseDate(row[state.dateColumn]);
    if (!dateValue) return false;
    if (start && dateValue < start) return false;
    if (end && dateValue > end) return false;
    return true;
  });
}

function renderKPIs(rows) {
  if (renderKPIs.running) {
    console.warn("renderKPIs re-entry detected");
    return;
  }
  renderKPIs.running = true;
  renderKpiCallCount += 1;
  console.log(`renderKPIs call #${renderKpiCallCount}`, { rows: rows.length, metric: state.selections.primaryMetric });
  kpiGrid.innerHTML = "";
  const chosenMetrics = chooseKpiMetrics(state.schema.profiles, state.schema.numeric);
  const metrics = (chosenMetrics && chosenMetrics.length ? chosenMetrics : (state.schema.numeric || []).slice(0, 6)) || [];

  (metrics || []).forEach((metric) => {
    const metricLabel = formatMetricLabel(metric);
    const values = rows
      .map((row) => parseNumber(row[metric]))
      .filter((value) => value !== null);
    const total = values.reduce((sum, value) => sum + value, 0);
    const avg = values.length ? total / values.length : 0;
    const median = computeMedian(values);
    const min = values.length ? Math.min(...values) : 0;
    const max = values.length ? Math.max(...values) : 0;

    const metricType = inferMetricType(metric, values);
    const primary = metricType.kind === "rate" ? avg : total;
    const change = computePeriodChange(rows, metric, metricType);
    const supportsTrend = Boolean(state.datasetCapabilities?.hasDateField && state.dateColumn);
    const miniTrend = supportsTrend ? buildUserKpiSparklineData(rows, metric, metricType) : null;
    const deltaTone = supportsTrend && change ? inferUserKpiDeltaTone(change.value) : "flat";
    const deltaBadge = supportsTrend && change
      ? `<em class="kpi-delta ${deltaTone}">${formatUserDeltaArrow(deltaTone)} ${escapeHtml(change.value)}</em>`
      : "";
    const sparkline = supportsTrend && miniTrend
      ? `<div class="kpi-mini-trend">${buildUserKpiSparklineSvg(miniTrend, metricType, metricLabel)}</div>
         <div class="kpi-sparkline-caption">Daily trend last 30 days compared to prior 30 days</div>`
      : `<div class="kpi-sparkline-caption">Snapshot metric, no time trend available.</div>`;

    const card = document.createElement("div");
    card.className = "kpi-card user-kpi-card";
    card.innerHTML = `
      <h4>${metricLabel}</h4>
      <div class="kpi-value">${formatMetricValue(primary, metricType)}</div>
      ${deltaBadge}
      ${sparkline}
      <div class="kpi-caption">${supportsTrend ? "Last 30 days vs prior 30 days" : "Snapshot metric, no time trend available."}</div>
      <div class="kpi-meta">
        <div>Avg: ${formatMetricValue(avg, metricType, true)}</div>
        <div>Min/Max: ${formatMetricValue(min, metricType, true)} · ${formatMetricValue(max, metricType, true)}</div>
      </div>
    `;
    kpiGrid.appendChild(card);
  });
  renderKPIs.running = false;
}
renderKPIs.running = false;

function inferUserKpiDeltaTone(changeValue) {
  const text = String(changeValue || "").trim();
  if (!text) return "flat";
  if (text.startsWith("-")) return "down";
  if (text === "0%" || text === "0.0 pp" || text === "0.0%") return "flat";
  return "up";
}

function formatUserDeltaArrow(tone) {
  if (tone === "up") return "▲";
  if (tone === "down") return "▼";
  return "•";
}

function buildUserKpiSparklineData(rows, metric, metricType) {
  if (!state.dateColumn) return null;
  const fullSeries = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange, metricType);
  const labels = fullSeries.labels || [];
  const values = fullSeries.values || [];
  if (!labels.length || !values.length) return null;
  const window = 30;
  const currentValues = values.slice(-window);
  const priorValues = values.slice(-window * 2, -window);
  const currentLabels = labels.slice(-window);
  const priorLabels = labels.slice(-window * 2, -window);
  const padSeries = (arr, size) => (arr.length >= size ? arr : [...Array(size - arr.length).fill(0), ...arr]);
  const padLabels = (arr, size) => (arr.length >= size ? arr : [...Array(size - arr.length).fill(""), ...arr]);
  return {
    currentSeries: padSeries(currentValues, window),
    priorSeries: padSeries(priorValues, window),
    currentLabels: padLabels(currentLabels, window),
    priorLabels: padLabels(priorLabels, window),
  };
}

function buildUserKpiSparklineSvg(series, metricType, label) {
  const currentSeries = series?.currentSeries || [];
  const priorSeries = series?.priorSeries || [];
  const currentLabels = series?.currentLabels || [];
  const priorLabels = series?.priorLabels || [];
  const length = Math.max(currentSeries.length, priorSeries.length, 30);
  const width = 240;
  const height = 52;
  const padX = 4;
  const padY = 4;
  const values = [...currentSeries, ...priorSeries];
  const minY = Math.min(...values, 0);
  const maxY = Math.max(...values, 0);
  const scaleX = (index) => padX + (index / Math.max(length - 1, 1)) * (width - padX * 2);
  const scaleY = (value) => (height - padY) - ((value - minY) / Math.max(maxY - minY, 1)) * (height - padY * 2);
  const pathFor = (arr) => arr.map((v, i) => `${i === 0 ? "M" : "L"}${scaleX(i)},${scaleY(v)}`).join(" ");
  const hoverRects = Array.from({ length }).map((_, i) => {
    const x = scaleX(i);
    const currentValue = currentSeries[i] ?? 0;
    const priorValue = priorSeries[i] ?? 0;
    const dateText = currentLabels[i] ? formatDateFriendly(currentLabels[i]) : `Day ${i + 1}`;
    const priorDateText = priorLabels[i] ? formatDateFriendly(priorLabels[i]) : `Day ${i + 1}`;
    const tooltip = `${dateText}\nCurrent: ${formatMetricValue(currentValue, metricType)}\nPrior (${priorDateText}): ${formatMetricValue(priorValue, metricType)}`;
    return `<rect x="${Math.max(0, x - 4)}" y="0" width="8" height="${height}" fill="transparent"><title>${tooltip}</title></rect>`;
  }).join("");
  return `
    <svg class="kpi-sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-label="${escapeHtml(label)} daily comparison sparkline">
      <path d="${pathFor(priorSeries)}" fill="none" stroke="rgba(161,161,170,0.55)" stroke-width="1.4" stroke-dasharray="3 2" />
      <path d="${pathFor(currentSeries)}" fill="none" stroke="rgba(96,165,250,0.95)" stroke-width="1.9" />
      ${hoverRects}
    </svg>
  `;
}

function formatDateFriendly(value) {
  const parsed = parseDate(value);
  if (!parsed) return String(value || "");
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function renderCharts(rows) {
  if (trendChartInstance) trendChartInstance.destroy();
  if (barChartInstance) barChartInstance.destroy();
  if (extraChartInstance) extraChartInstance.destroy();

  const metric = state.selections.primaryMetric || chooseTopNumericByVariance(state.schema.profiles, state.schema.numeric || []);
  state.selections.primaryMetric = metric;
  state.selectedMetric = metric;
  const metricLabel = formatMetricLabel(metric);
  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);

  if (state.dateColumn) {
    state.chosenXAxisType = "date";
    const bucket = rows.length > 200 ? "week" : "day";
    const series = aggregateByDate(rows, state.dateColumn, metric, bucket, state.dateRange, metricType);
    trendTitle.textContent = "Month over month trend";
    const periodLabel = series.labels.length < 3 ? `${series.labels.length} periods (limited)` : `${series.labels.length} periods`;
    trendSubtitle.textContent = `Metric: ${metricLabel} · ${periodLabel}`;
    trendChartInstance = createLineChart("trendChart", series.labels, series.values, metricType, series.counts);
  } else {
    const category = chooseCategoryColumn(rows, state.schema.profiles);
    if (category) {
      state.chosenXAxisType = "category";
      const series = aggregateByCategory(rows, category, metric, state.topN, metricType);
      trendTitle.textContent = "Month over month trend";
      trendSubtitle.textContent = "No date field detected in this dataset. Time based trends are unavailable.";
      trendChartInstance = createLineChart("trendChart", series.labels, series.values, metricType, series.counts);
    } else {
      state.chosenXAxisType = "index";
      const labels = rows.map((_, index) => index + 1);
      const values = rows.map((row) => parseNumber(row[metric])).map((value) => value ?? 0);
      trendTitle.textContent = "Month over month trend";
      trendSubtitle.textContent = "No date field detected in this dataset. Time based trends are unavailable.";
      trendChartInstance = createLineChart("trendChart", labels, values, metricType);
    }
  }

  let dimension = state.selectedDimension || chooseBestDimension(state.schema.profiles, state.schema.categoricals || []) || state.dateColumn;
  let breakdown = null;
  if (dimension) {
    state.selectedDimension = dimension;
    if (dimensionSelect.value !== dimension) dimensionSelect.value = dimension;
    breakdown = aggregateByCategory(rows, dimension, metric, state.topN, metricType);
    barTitle.textContent = "Performance details";
    barSubtitle.textContent = `Dimension: ${dimension}`;
    barChartInstance = createBarChart("barChart", breakdown.labels, breakdown.values, metricType, breakdown.counts);
  } else {
    state.selectedDimension = null;
    if (dimensionSelect) dimensionSelect.value = "";
    const labels = rows.map((_, index) => index + 1);
    const values = rows.map((row) => parseNumber(row[metric]) ?? 0);
    barTitle.textContent = "Performance details";
    barSubtitle.textContent = "No category column detected";
    barChartInstance = createBarChart("barChart", labels, values, metricType);
  }

  const correlationPair = findStrongestCorrelation(rows, state.schema.numeric);
  if (correlationPair) {
    extraTitle.textContent = `Correlation: ${correlationPair.x} vs ${correlationPair.y}`;
    extraSubtitle.textContent = `r = ${correlationPair.corr.toFixed(2)}`;
    extraChartInstance = createScatterChart("extraChart", correlationPair.points);
  } else {
    extraTitle.textContent = `Distribution of ${metric}`;
    extraSubtitle.textContent = "Histogram";
    extraChartInstance = createHistogram("extraChart", metricValues, metricType);
  }
}

function renderHeatmap() {
  return;
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
  Object.entries(profile || {}).forEach(([col, stats]) => {
    const tr = document.createElement("tr");
    const dateRange = stats.dateMin && stats.dateMax
      ? `${formatDateInput(stats.dateMin)} → ${formatDateInput(stats.dateMax)}`
      : "—";
    const topValues = stats.topValues?.length
      ? stats.topValues.map((item) => `${item.value} (${item.count})`).join(", ")
      : "—";

    const metricType = stats.type === "numeric" ? inferMetricType(col, []) : { isCurrency: false, isRate: false, kind: "generic" };
    const profileKind = metricType.kind === "rate" ? "generic" : metricType.kind;
    const cells = [
      col,
      stats.type,
      `${Math.round(stats.missingRate * 100)}%`,
      stats.uniqueCount ?? "—",
      stats.min !== null ? formatMetric(stats.min, profileKind) : "—",
      stats.max !== null ? formatMetric(stats.max, profileKind) : "—",
      stats.mean !== null ? formatMetric(stats.mean, profileKind) : "—",
      stats.median !== null ? formatMetric(stats.median, profileKind) : "—",
      stats.std !== null ? numberFormatter.format(stats.std) : "—",
      stats.type === "date" ? dateRange : "—",
      stats.type === "categorical" ? topValues : "—",
    ];

    (cells || []).forEach((value) => {
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
    qualityBadge.style.background = "rgba(124,58,237,0.18)";
    qualityBadge.style.color = "#ffffff";
    qualityBadge.style.borderColor = "rgba(124,58,237,0.45)";
  } else if (score >= 70) {
    qualityBadge.style.background = "rgba(139,92,246,0.16)";
    qualityBadge.style.color = "#e9d5ff";
    qualityBadge.style.borderColor = "rgba(139,92,246,0.4)";
  } else {
    qualityBadge.style.background = "rgba(124,58,237,0.10)";
    qualityBadge.style.color = "#e5e7eb";
    qualityBadge.style.borderColor = "rgba(124,58,237,0.25)";
  }
}

function computeQualityScore(rows) {
  let score = 100;
  const profile = state.schema.profiles;
  const missingPenalty = Object.values(profile).reduce((sum, stats) => sum + stats.missingRate, 0) / state.schema.columns.length;
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
  (columns || []).forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.addEventListener("click", () => sortTable(col));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  (currentTableRows || []).forEach((row) => {
    const tr = document.createElement("tr");
    (columns || []).forEach((col) => {
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
  const lower = String(metric || "").toLowerCase();
  const isCurrency = /(revenue|sales|cost|spend|profit|price|amount|gmv|mrr|arr)/i.test(lower);
  const isRateName = /(rate|percent|ctr|cvr|roas|ratio|retention)/i.test(lower);
  const isDuration = /(time|duration|seconds|secs|minutes|mins|hours)/i.test(lower);
  const isRateRange = values.length > 0 && values.every((val) => val >= 0 && val <= 1);
  const isRate = isRateName || isRateRange;
  const kind = isRate ? "rate" : isCurrency ? "currency" : isDuration ? "duration" : "count";
  return { isCurrency, isRate, isDuration, kind };
}

function computeRateValue(metric, values) {
  if (!values.length) return 0;
  const denominator = findRateDenominator(metric, state.schema.numeric || []);
  if (!denominator) {
    const avg = values.reduce((sum, val) => sum + val, 0) / values.length;
    return avg;
  }
  const numerator = values.reduce((sum, val) => sum + val, 0);
  const denomValues = state.filteredRows.map((row) => parseNumber(row[denominator])).filter((value) => value !== null);
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
        const found = (numericColumns || []).find((col) => denomPattern.test(col.toLowerCase()));
        if (found) return found;
      }
    }
  }
  return null;
}

function formatMetric(value, metricKind, decimals = 2, compact = false) {
  if (value === null || value === undefined || Number.isNaN(value)) return "—";
  if (metricKind === "rate") {
    const pct = value * 100;
    return `${pct.toFixed(1)}%`;
  }
  if (metricKind === "currency") {
    return currencyFormatter.format(value);
  }
  if (metricKind === "duration") {
    return `${Number(value).toFixed(1)}s`;
  }
  const formatter = compact ? new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: decimals }) : numberFormatter;
  return formatter.format(value);
}

function formatMetricValue(value, metricType, compact = false) {
  const kind = metricType?.kind || (metricType?.isRate ? "rate" : metricType?.isCurrency ? "currency" : "count");
  return formatMetric(value, kind, 2, compact);
}

function computeMedian(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function computeCoefficientOfVariation(rows, metric) {
  const values = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  if (values.length < 2) return null;
  const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
  if (mean === 0) return null;
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  const std = Math.sqrt(variance);
  return std / Math.abs(mean);
}

function computePeriodChange(rows, metric, metricType) {
  if (!state.dateColumn) return null;
  const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange, metricType);
  if (series.values.length < 6) return null;
  const window = series.values.length >= 14 ? 7 : 3;
  const recent = series.values.slice(-window);
  const previous = series.values.slice(-window * 2, -window);
  if (!previous.length) return null;
  const recentTotal = recent.reduce((sum, val) => sum + val, 0) / recent.length;
  const previousTotal = previous.reduce((sum, val) => sum + val, 0) / previous.length;
  const label = `Last ${window} vs prior ${window}`;
  if (metricType?.kind === "rate") {
    const delta = recentTotal - previousTotal;
    const value = `${(delta * 100).toFixed(1)} pp`;
    return { label, value };
  }
  const change = previousTotal === 0 ? 0 : (recentTotal - previousTotal) / previousTotal;
  const value = percentFormatter.format(change);
  return { label, value };
}

function aggregateByDate(rows, dateColumn, metric, bucket, dateRange, metricType) {
  const totals = new Map();
  const counts = new Map();
  (rows || []).forEach((row) => {
    const dateValue = parseDate(row[dateColumn]);
    const metricValue = parseNumber(row[metric]);
    if (!dateValue || metricValue === null) return;
    if (dateRange.start && dateValue < dateRange.start) return;
    if (dateRange.end && dateValue > dateRange.end) return;
    const key = bucket === "week" ? getWeekStart(dateValue) : formatDateInput(dateValue);
    totals.set(key, (totals.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const labels = Array.from(totals.keys()).sort();
  const values = labels.map((label) => {
    const total = totals.get(label) || 0;
    const count = counts.get(label) || 1;
    if (metricType?.kind === "rate") return total / count;
    return total;
  });
  const ns = labels.map((label) => counts.get(label) || 0);
  return { labels, values, counts: ns };
}

function aggregateByCategory(rows, dimension, metric, topN, metricType) {
  if (!dimension) {
    return { labels: [], values: [], counts: [] };
  }
  const totals = new Map();
  const counts = new Map();
  (rows || []).forEach((row) => {
    const key = row[dimension] ? String(row[dimension]).trim() : "Unknown";
    const metricValue = parseNumber(row[metric]);
    if (metricValue === null) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const entries = Array.from(totals.entries())
    .map(([label, total]) => ({
      label,
      value: metricType?.kind === "rate" ? total / (counts.get(label) || 1) : total,
      count: counts.get(label) || 0,
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, topN);

  return {
    labels: entries.map((item) => item.label),
    values: entries.map((item) => item.value),
    counts: entries.map((item) => item.count),
  };
}
function createLineChart(canvasId, labels, values, metricType, counts = []) {
  const ctx = document.getElementById(canvasId);
  return new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: values,
          borderColor: "#7c3aed",
          backgroundColor: "rgba(124, 58, 237, 0.2)",
          tension: 0.35,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 5,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => {
              const n = counts?.[context.dataIndex] ?? null;
              const valueLabel = formatMetricValue(context.parsed.y, metricType);
              return n ? `${valueLabel} (n=${n})` : valueLabel;
            },
          },
        },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: metricType?.kind === "rate"
          ? {
              min: 0,
              max: 1,
              ticks: {
                callback: (value) => `${(Number(value) * 100).toFixed(0)}%`,
              },
            }
          : { beginAtZero: true },
      },
    },
  });
}

function createBarChart(canvasId, labels, values, metricType, counts = []) {
  const ctx = document.getElementById(canvasId);
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: values,
          backgroundColor: "#8b5cf6",
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
            label: (context) => {
              const n = counts?.[context.dataIndex] ?? null;
              const valueLabel = formatMetricValue(context.parsed.y, metricType);
              return n ? `${valueLabel} (n=${n})` : valueLabel;
            },
          },
        },
        datalabels: {
          color: "#ffffff",
          anchor: "end",
          align: "end",
          formatter: (value) => formatMetricValue(value, metricType),
          clamp: true,
        },
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6 } },
        y: metricType?.kind === "rate"
          ? {
              min: 0,
              max: 1,
              ticks: {
                callback: (value) => `${(Number(value) * 100).toFixed(0)}%`,
              },
            }
          : { beginAtZero: true },
      },
    },
    plugins: [ChartDataLabels],
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
          backgroundColor: "rgba(124, 58, 237, 0.5)",
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

function createHistogram(canvasId, values, metricType) {
  const ctx = document.getElementById(canvasId);
  const bins = 10;
  if (!values.length) {
    return new Chart(ctx, { type: "bar", data: { labels: [], datasets: [] } });
  }
  const min = metricType?.kind === "rate" ? 0 : Math.min(...values);
  const max = metricType?.kind === "rate" ? 1 : Math.max(...values);
  const step = (max - min) / bins || 1;
  const counts = Array.from({ length: bins }, () => 0);
  values.forEach((value) => {
    if (metricType?.kind === "rate" && (value < 0 || value > 1)) return;
    const index = Math.min(Math.floor((value - min) / step), bins - 1);
    counts[index] += 1;
  });
  const labels = counts.map((_, i) => {
    const start = min + i * step;
    const end = min + (i + 1) * step;
    if (metricType?.kind === "rate") {
      return `${(start * 100).toFixed(0)}-${(end * 100).toFixed(0)}%`;
    }
    return `${start.toFixed(1)}-${end.toFixed(1)}`;
  });
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          data: counts,
          backgroundColor: "#8b5cf6",
        },
      ],
    },
    options: { plugins: { legend: { display: false } } },
  });
}

function buildCorrelationMatrix() {
  return null;
}

function findStrongestCorrelation(rows, numericColumns) {
  if (!Array.isArray(numericColumns) || numericColumns.length < 2) return null;
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
  const pairs = (rows || [])
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
  const metric = state.selections.primaryMetric;
  if (metric) {
    const missingRate = state.schema.profiles[metric]?.missingRate || 0;
    if (missingRate > 0.3) {
      warnings.push(`Selected metric "${metric}" has ${Math.round(missingRate * 100)}% missing values.`);
    }
    const nonNumericRate = state.schema.profiles[metric]?.nonNumericRate || 0;
    if (nonNumericRate > 0.05) {
      warnings.push(`Metric "${metric}" has ${Math.round(nonNumericRate * 100)}% non-numeric values.`);
    }
  }

  if (state.selectedDimension) {
    const uniqueCount = state.schema.profiles[state.selectedDimension]?.uniqueCount || 0;
    if (uniqueCount > 20) {
      warnings.push(`"${state.selectedDimension}" has high cardinality (${uniqueCount} unique values).`);
    }
  }

  const duplicateCount = countDuplicateRows(state.filteredRows);
  if (duplicateCount > 0) {
    warnings.push(`${duplicateCount} duplicate rows detected.`);
  }

  if (state.dateColumn) {
    const gaps = countDateGaps(state.filteredRows, state.dateColumn);
    if (gaps > 0) {
      warnings.push(`Detected ${gaps} missing date buckets in the trend.`);
    }
  }

  return warnings;
}

function countDuplicateRows(rows) {
  if (!Array.isArray(rows)) return 0;
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
  if (!Array.isArray(rows)) return 0;
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
  const evidencePanel = document.getElementById("userEvidencePanel");
  if (evidencePanel) {
    evidencePanel.classList.add("hidden");
    evidencePanel.innerHTML = `<div class="helper-text">Select an insight to view evidence details.</div>`;
  }
  const insights = buildUnifiedInsights(rows);
  const insight = insights[0];
  if (!insight) return;
  const card = document.createElement("li");
  card.className = "insight-card user-insight-card user-insight-summary-card";
  const confidenceLevel = insight?.confidence?.level || inferUserInsightConfidence(insight);
  const confidenceLabel = confidenceLevel.charAt(0).toUpperCase() + confidenceLevel.slice(1);
  const whyLine = confidenceLevel !== "high"
    ? `<p class="insight-meta confidence-why"><strong>Why this confidence:</strong> ${escapeHtml((insight?.confidence?.reasons || [buildUserInsightConfidenceReason(insight)]).slice(0, 2).join(" "))}</p>`
    : "";
  const metricQuality = classifyNumericMetricColumn(
    state.selections.primaryMetric,
    state.schema.profiles[state.selections.primaryMetric],
    rows.length
  );
  const warningTag = metricQuality.isIdentifierLike
    ? `<span class="metric-warning-tag">Identifier-like metric, insights may be low quality</span>`
    : "";
  const diagnostics = insight?.diagnostics || {};
  const diagnosticsMarkup = `
    <details class="insight-diagnostics">
      <summary>More diagnostics</summary>
      <div class="diagnostics-grid">
        <div>
          <strong>Anomalies</strong>
          <ul>${(diagnostics.anomalies || ["No notable anomalies detected."]).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        </div>
        <div>
          <strong>Concentration risk</strong>
          <ul>${(diagnostics.concentrationRisk || ["Concentration risk appears moderate."]).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        </div>
        <div>
          <strong>Drivers and relationships</strong>
          <ul>${(diagnostics.relationships || ["No additional relationships identified."]).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        </div>
      </div>
    </details>
  `;
  card.innerHTML = `
    <div class="insight-summary-header">
      <h4>${escapeHtml(insight.headline || insight.title || "Executive summary")}</h4>
      ${warningTag}
    </div>
    <section class="insight-section">
      <h5>Executive summary</h5>
      <p class="insight-meta">${escapeHtml(insight.executiveSummary || insight.deltaSummary || insight.summary || "")}</p>
      <ul class="insight-mini-list">
        ${(insight.executiveBullets || []).slice(0, 2).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
      </ul>
    </section>
    <section class="insight-section">
      <h5>Top driver</h5>
      <p class="insight-meta">${escapeHtml(insight.driverNarrative || "No dominant driver identified.")}</p>
      <p class="insight-meta subtle">${escapeHtml(insight.deltaSummary || "")}</p>
    </section>
    <section class="insight-section">
      <h5>Recommended action</h5>
      <p class="insight-meta">${escapeHtml(insight.action || "Investigate top drivers and validate with a controlled test.")}</p>
      <span class="severity ${confidenceLevel} confidence-pill" tabindex="0" aria-label="Confidence info: Confidence reflects data coverage, consistency and evidence quality for this insight.">
        Confidence: ${confidenceLabel}
        <span class="confidence-tooltip" role="tooltip">Confidence reflects data coverage, sample size, and consistency over time.</span>
      </span>
      ${whyLine}
    </section>
    ${diagnosticsMarkup}
    <button type="button" class="ghost view-evidence">View evidence</button>
  `;
  const button = card.querySelector(".view-evidence");
  if (button) {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      renderUserInsightEvidence(evidencePanel, insight);
    });
  }
  insightsList.appendChild(card);
}

function extractInsightBullet(insight, prefix) {
  return (insight?.bullets || []).find((bullet) => String(bullet).startsWith(prefix)) || "";
}

function inferUserInsightConfidence(insight) {
  const bullets = (insight?.bullets || []).join(" ").toLowerCase();
  if (bullets.includes("insufficient") || bullets.includes("skipped") || bullets.includes("fewer than")) return "low";
  if (bullets.includes("limited") || bullets.includes("no date column") || bullets.includes("validate")) return "medium";
  return "high";
}

function buildUserInsightConfidenceReason(insight) {
  const bullets = (insight?.bullets || []).join(" ");
  if (/insufficient|skipped|fewer than/i.test(bullets)) {
    return "Limited supporting data or validation coverage for this insight.";
  }
  if (/no date column|limited|validate/i.test(bullets)) {
    return "Evidence is directionally useful, but trend or segment validation is limited.";
  }
  return "Evidence quality is moderate and should be validated with additional slices.";
}

function renderUserInsightEvidence(panel, insight) {
  if (!panel) return;
  panel.classList.remove("hidden");
  if (!insight?.evidence) {
    const bullets = insight?.bullets || [];
    panel.innerHTML = `
      <div class="evidence-section">
        <strong>Evidence detail</strong>
        <div class="helper-text">${escapeHtml(insight.summary || "")}</div>
      </div>
      <div class="evidence-section">
        <strong>Analytical proof</strong>
        <div class="evidence-table-wrap">
          <table class="evidence-table">
            <thead>
              <tr><th>Type</th><th>Detail</th></tr>
            </thead>
            <tbody>
              ${bullets.map((bullet) => {
                const parts = String(bullet).split(":");
                const type = parts.length > 1 ? parts.shift() : "Note";
                const detail = parts.length > 0 ? parts.join(":").trim() : String(bullet);
                return `<tr><td>${escapeHtml(String(type))}</td><td>${escapeHtml(detail)}</td></tr>`;
              }).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;
    return;
  }
  const evidence = insight.evidence;
  const comparisonRows = (evidence.comparisonTable || []).map((row) => `
    <tr>
      <td>${escapeHtml(String(row.segment))}</td>
      <td>${escapeHtml(String(row.value))}</td>
      <td>${escapeHtml(String(row.share))}</td>
      <td>${escapeHtml(String(row.rank))}</td>
    </tr>
  `).join("");
  const driverRows = (evidence.driverBreakdownTable || []).map((row) => `
    <tr>
      <td>${escapeHtml(String(row.combination))}</td>
      <td>${escapeHtml(String(row.current))}</td>
      <td>${escapeHtml(String(row.prior))}</td>
      <td>${escapeHtml(String(row.delta))}</td>
      <td>${escapeHtml(String(row.liftShare))}</td>
    </tr>
  `).join("");
  const confidenceReasons = (insight?.confidence?.reasons || []).map((reason) => `<li>${escapeHtml(reason)}</li>`).join("");
  const periods = insight?.confidence?.metrics?.periods ?? 0;
  const periodsText = periods === 1
    ? "Only one time period available. Trend reliability is limited."
    : periods > 1 && periods < 3
      ? `Limited history: this insight is based on only ${periods} time period(s).`
      : periods >= 3
        ? `Based on ${periods} time periods.`
        : "No historical time buckets available.";
  const trendConsistency = evidence.trendSummary
    ? `
      <div class="evidence-section">
        <strong>Trend consistency</strong>
        <div class="evidence-stats">
          <div><span>Periods analysed</span><strong>${escapeHtml(String(evidence.trendSummary.periodsAnalysed ?? periods))}</strong><p class="helper-text">${escapeHtml(periodsText)}</p></div>
          <div><span>Stability score</span><strong>${escapeHtml(String(evidence.consistencyScore ?? "—"))}</strong></div>
          <div><span>Variance</span><strong>${escapeHtml(String(evidence.trendSummary.variance ?? "—"))}</strong></div>
        </div>
      </div>`
    : `
      <div class="evidence-section">
        <strong>Trend consistency</strong>
        <div class="helper-text">Trend analysis unavailable due to missing date field.</div>
      </div>`;

  panel.innerHTML = `
    <div class="evidence-section">
      <strong>Segment comparison</strong>
      <div class="evidence-table-wrap">
        <table class="evidence-table">
          <thead><tr><th>Segment</th><th>Value</th><th>Share</th><th>Rank</th></tr></thead>
          <tbody>${comparisonRows || `<tr><td colspan="4">No segment comparison available.</td></tr>`}</tbody>
        </table>
      </div>
    </div>
    <div class="evidence-section">
      <strong>What drives this insight</strong>
      <div class="helper-text">${escapeHtml(evidence.contributionSummary || evidence.driverNarrative || "No dominant driver identified.")}</div>
    </div>
    <div class="evidence-section">
      <strong>Driver breakdown</strong>
      <div class="evidence-table-wrap">
        <table class="evidence-table">
          <thead><tr><th>Combination</th><th>Current</th><th>Prior</th><th>Delta</th><th>Lift share</th></tr></thead>
          <tbody>${driverRows || `<tr><td colspan="5">No dominant driver combination identified.</td></tr>`}</tbody>
        </table>
      </div>
      <div class="helper-text">${escapeHtml((evidence.filtersUsed || []).join(" · "))}</div>
    </div>
    ${trendConsistency}
    <div class="evidence-section">
      <strong>Confidence explanation</strong>
      ${confidenceReasons ? `<ul class="evidence-reasons">${confidenceReasons}</ul>` : `<div class="helper-text">No confidence factors available.</div>`}
    </div>
  `;
}

function buildUnifiedInsights(rows) {
  const metric = state.selections.primaryMetric;
  if (!metric) return buildDecisionInsights(rows);
  if (metric === "__row_count__" || classifyNumericMetricColumn(metric, state.schema.profiles[metric], rows.length).isIdentifierLike) {
    return [buildSafeDescriptiveInsight(rows, metric)];
  }
  const dimension = state.selectedDimension || chooseBestDimension(state.schema.profiles, state.schema.categoricals || []);
  const metricValues = (rows || []).map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  if (!dimension) {
    return buildDecisionInsights(rows);
  }
  const comparison = aggregateByCategory(rows, dimension, metric, Math.max(state.topN || 8, 8), metricType);
  if ((comparison.labels || []).length < 2) {
    return buildDecisionInsights(rows);
  }

  const totalValue = (comparison.values || []).reduce((sum, value) => sum + (Number(value) || 0), 0);
  const comparisonTable = comparison.labels.map((label, index) => ({
    segment: label,
    value: formatMetricValue(comparison.values[index], metricType),
    share: totalValue ? percentFormatter.format((comparison.values[index] || 0) / totalValue) : "—",
    rank: index + 1,
    rawValue: comparison.values[index] || 0,
  }));

  const top = comparisonTable[0];
  const runnerUp = comparisonTable[1];
  const deltaRaw = (top?.rawValue || 0) - (runnerUp?.rawValue || 0);
  const deltaPercent = (runnerUp?.rawValue || 0) > 0 ? (deltaRaw / runnerUp.rawValue) * 100 : 0;
  const deltaSummary = metricType.kind === "rate"
    ? `Delta ${((deltaRaw || 0) * 100).toFixed(1)} pp (${top.segment} vs ${runnerUp.segment})`
    : `Delta ${formatMetricValue(deltaRaw, metricType)} (${top.segment} vs ${runnerUp.segment})`;

  const periodSplit = splitRowsIntoCurrentPrior(rows);
  const driverInfo = computeUserTopDriver({
    rows,
    metric,
    metricType,
    primaryDimension: dimension,
    topSegment: top.segment,
    periodSplit,
  });
  const periods = periodSplit.hasDateField ? periodSplit.periodsAnalysed : 0;
  const confidence = computeUserUnifiedConfidence({
    periods,
    sampleSize: driverInfo.sampleSize || comparison.counts?.[0] || 0,
    deltaPercent: Math.abs(deltaPercent || 0),
    variance: driverInfo.varianceScore ?? 0,
    hasDateField: periodSplit.hasDateField,
  });
  const action = buildUserDynamicAction({
    metricKey: metric,
    topSegment: top.segment,
    driverCombination: driverInfo.driverCombinationText || top.segment,
    confidence,
    deltaStrength: Math.abs(deltaPercent || 0),
  });

  const insight = {
    id: `user:${metric}:${dimension}:${top.segment}`,
    headline: `${formatMetricLabel(metric)} is higher in ${top.segment} than ${runnerUp.segment}.`,
    deltaSummary,
    metricKey: metric,
    dimensionKey: dimension,
    topSegment: top.segment,
    driverCombination: driverInfo.driverCombinationText || top.segment,
    driverNarrative: driverInfo.narrative,
    confidence,
    action,
    evidence: {
      comparisonTable: comparisonTable.slice(0, 8),
      contributionSummary: driverInfo.contributionSummary,
      trendSummary: periodSplit.hasDateField ? {
        periodsAnalysed: periods,
        variance: driverInfo.varianceScore == null ? "—" : driverInfo.varianceScore.toFixed(2),
      } : null,
      consistencyScore: periodSplit.hasDateField ? Math.max(1, Math.min(100, Math.round((1 - Math.min(driverInfo.varianceScore || 0, 1)) * 100))) : null,
      driverBreakdownTable: driverInfo.breakdownTable,
      driverNarrative: driverInfo.narrative,
      filtersUsed: [
        `Metric: ${formatMetricLabel(metric)}`,
        `Dimension: ${dimension}`,
        periodSplit.hasDateField ? "Window: Last 30 days vs prior 30 days" : "Window: Snapshot (no date field)",
      ],
    },
    executiveSummary: `${formatMetricLabel(metric)} shows a clear gap between ${top.segment} and ${runnerUp.segment}${metricQualitySuffix(metric, rows.length)}`,
    executiveBullets: [
      `${top.segment} ranks #1 with ${top.value}; ${runnerUp.segment} ranks #2 with ${runnerUp.value}.`,
      `Top driver: ${driverInfo.driverCombinationText || top.segment} (${percentFormatter.format((driverInfo.topLiftShare || 0))} of measured lift).`,
    ],
    diagnostics: buildInsightDiagnostics({
      rows,
      metric,
      metricType,
      comparisonTable,
      driverInfo,
      deltaPercent,
      dimension,
    }),
  };
  return [insight];
}

function buildSafeDescriptiveInsight(rows, metric) {
  const dimension = state.selectedDimension || chooseBestDimension(state.schema.profiles, state.schema.categoricals || []) || "segment";
  const hasDate = Boolean(state.dateColumn && state.datasetCapabilities?.hasDateField);
  const confidence = {
    level: hasDate ? "low" : "medium",
    reasons: [
      metric === "__row_count__"
        ? "Limited structured metrics detected. Insights are descriptive only."
        : "Selected metric behaves like an identifier and is not suitable for causal insights.",
      hasDate ? "Not enough reliable metric semantics for trend-based claims." : "No historical time buckets available.",
    ],
    metrics: { periods: hasDate ? 1 : 0, sampleSize: rows.length, deltaPercent: 0, variance: 0 },
  };
  return {
    id: `safe:${metric}:${dimension}`,
    headline: "Structured tables detected, but insight quality is limited.",
    title: "Structured tables detected, but insight quality is limited.",
    executiveSummary: "This file contains structured tables but time trends or metric semantics are limited. Review top segments by spend and conversions where available.",
    executiveBullets: [
      `Use ${dimension} to review top segments by core marketing metrics.`,
      "Prefer Spend, Clicks, Impressions, Conversions, CTR, CPC, CPA, or eCPM as the primary metric.",
    ],
    deltaSummary: "Descriptive mode only; trend claims are suppressed.",
    driverNarrative: "No dominant driver is shown because the selected metric is not reliable for decisioning.",
    action: "Select a marketing metric such as Spend, Conversions, Clicks, or Impressions to generate higher-confidence insights.",
    confidence,
    diagnostics: {
      anomalies: ["Trend-based insights are blocked until a meaningful metric is selected."],
      concentrationRisk: ["Review top segments manually in Performance details."],
      relationships: ["Correlation analysis is less reliable on identifier-like or row count metrics."],
    },
    evidence: {
      comparisonTable: [],
      contributionSummary: "Evidence is withheld because the selected metric is not suitable for analytical claims.",
      trendSummary: hasDate ? { periodsAnalysed: 1, variance: "—" } : null,
      consistencyScore: null,
      driverBreakdownTable: [],
      filtersUsed: [
        `Metric: ${formatMetricLabel(metric)}`,
        `Dimension: ${dimension}`,
        hasDate ? "Time context detected but not used for claims" : "No time buckets detected",
      ],
    },
  };
}

function metricQualitySuffix(metric, rowCount) {
  const quality = classifyNumericMetricColumn(metric, state.schema.profiles[metric], rowCount);
  return quality.isIdentifierLike ? ", but the selected metric behaves like an identifier." : ".";
}

function splitRowsIntoCurrentPrior(rows) {
  if (!state.dateColumn) {
    return { hasDateField: false, currentRows: rows || [], priorRows: [], periodsAnalysed: 0 };
  }
  const datedRows = (rows || [])
    .map((row) => ({ row, date: parseDate(row[state.dateColumn]) }))
    .filter((item) => item.date);
  if (!datedRows.length) {
    return { hasDateField: false, currentRows: rows || [], priorRows: [], periodsAnalysed: 0 };
  }
  const maxDate = new Date(Math.max(...datedRows.map((item) => item.date.getTime())));
  const currentStart = new Date(maxDate);
  currentStart.setDate(currentStart.getDate() - 29);
  const priorEnd = new Date(currentStart);
  priorEnd.setDate(priorEnd.getDate() - 1);
  const priorStart = new Date(priorEnd);
  priorStart.setDate(priorStart.getDate() - 29);
  const currentRows = [];
  const priorRows = [];
  datedRows.forEach((item) => {
    if (item.date >= currentStart && item.date <= maxDate) currentRows.push(item.row);
    else if (item.date >= priorStart && item.date <= priorEnd) priorRows.push(item.row);
  });
  const dayKeys = new Set(datedRows.map((item) => formatDateInput(item.date)));
  return {
    hasDateField: true,
    currentRows,
    priorRows,
    periodsAnalysed: Math.min(12, dayKeys.size >= 365 ? 12 : dayKeys.size >= 90 ? 3 : dayKeys.size >= 30 ? 2 : 1),
    currentRange: [currentStart, maxDate],
    priorRange: [priorStart, priorEnd],
  };
}

function aggregateMetricForRows(rows, metric, metricType) {
  const values = (rows || []).map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  if (!values.length) return { value: 0, count: 0 };
  if (metricType?.kind === "rate") {
    return { value: values.reduce((sum, value) => sum + value, 0) / values.length, count: values.length };
  }
  return { value: values.reduce((sum, value) => sum + value, 0), count: values.length };
}

function computeUserTopDriver({ rows, metric, metricType, primaryDimension, topSegment, periodSplit }) {
  const candidates = (state.schema.categoricals || []).filter((col) => col && col !== primaryDimension).slice(0, 4);
  const filteredRows = (rows || []).filter((row) => String(row[primaryDimension] ?? "Unknown").trim() === String(topSegment));
  const sourceRows = filteredRows.length ? filteredRows : (rows || []);
  const combos = [];
  const comboDefs = [];
  candidates.forEach((dim) => comboDefs.push([primaryDimension, dim]));
  if (candidates.length >= 2) comboDefs.push([primaryDimension, candidates[0], candidates[1]]);
  if (!comboDefs.length) comboDefs.push([primaryDimension]);

  comboDefs.forEach((dims) => {
    const map = new Map();
    sourceRows.forEach((row) => {
      const keyParts = dims.map((dim) => String(row[dim] ?? "Unknown").trim() || "Unknown");
      if (String(row[primaryDimension] ?? "Unknown").trim() !== String(topSegment)) return;
      const key = keyParts.join(" | ");
      if (!map.has(key)) {
        map.set(key, { key, dims, keyParts, rows: [] });
      }
      map.get(key).rows.push(row);
    });
    map.forEach((entry) => combos.push(entry));
  });

  if (!combos.length) {
    return {
      narrative: `No dominant driver identified.`,
      contributionSummary: `No dominant driver identified.`,
      driverCombinationText: topSegment,
      breakdownTable: [],
      sampleSize: 0,
      varianceScore: null,
    };
  }

  const scored = combos.map((combo) => {
    const current = periodSplit.hasDateField
      ? aggregateMetricForRows((periodSplit.currentRows || []).filter((row) => combo.rows.includes(row)), metric, metricType)
      : aggregateMetricForRows(combo.rows, metric, metricType);
    const prior = periodSplit.hasDateField
      ? aggregateMetricForRows((periodSplit.priorRows || []).filter((row) => combo.rows.includes(row)), metric, metricType)
      : { value: 0, count: 0 };
    const delta = periodSplit.hasDateField ? (current.value - prior.value) : current.value;
    return {
      ...combo,
      currentValue: current.value,
      priorValue: prior.value,
      delta,
      sampleSize: current.count + prior.count,
    };
  });

  const totalDelta = scored.reduce((sum, item) => sum + Math.max(item.delta, 0), 0) || 0;
  const winnerTotal = scored.reduce((sum, item) => sum + Math.max(item.currentValue, 0), 0) || 0;
  scored.forEach((item) => {
    item.shareOfWinner = winnerTotal > 0 ? item.currentValue / winnerTotal : 0;
    item.liftShare = totalDelta > 0 ? Math.max(item.delta, 0) / totalDelta : 0;
    item.score = (item.liftShare * 0.6) + (item.shareOfWinner * 0.4);
  });
  scored.sort((a, b) => b.score - a.score);
  const topDriver = scored[0];

  const comboLabel = formatDriverCombinationLabel(topDriver?.keyParts || []);
  const liftPct = percentFormatter.format(topDriver?.liftShare || 0);
  const narrative = topDriver && (topDriver.liftShare > 0.05 || topDriver.shareOfWinner > 0.2)
    ? `${comboLabel} contributes ${liftPct} of the lift.`
    : "No dominant driver identified.";
  const contributionSummary = topDriver && periodSplit.hasDateField
    ? `${comboLabel} contributes ${liftPct} of the lift versus the prior 30-day period.`
    : topDriver
      ? `${comboLabel} is the largest contributing combination in the current snapshot.`
      : "No dominant driver identified.";
  const breakdownTable = scored.slice(0, 5).map((item) => ({
    combination: formatDriverCombinationLabel(item.keyParts),
    current: formatMetricValue(item.currentValue, metricType),
    prior: periodSplit.hasDateField ? formatMetricValue(item.priorValue, metricType) : "—",
    delta: periodSplit.hasDateField
      ? (metricType.kind === "rate" ? `${((item.delta || 0) * 100).toFixed(1)} pp` : formatMetricValue(item.delta, metricType))
      : "—",
    liftShare: percentFormatter.format(item.liftShare || 0),
  }));

  let varianceScore = null;
  if (periodSplit.hasDateField && topDriver) {
    const daySeries = aggregateByDate(topDriver.rows, state.dateColumn, metric, "day", state.dateRange, metricType);
    const values = daySeries.values || [];
    if (values.length >= 2) {
      const mean = values.reduce((sum, v) => sum + v, 0) / values.length || 0;
      if (mean !== 0) {
        const variance = values.reduce((sum, v) => sum + ((v - mean) ** 2), 0) / values.length;
        varianceScore = Math.min(1, Math.sqrt(variance) / Math.abs(mean));
      } else {
        varianceScore = 1;
      }
    }
  }

  return {
    narrative,
    contributionSummary,
    driverCombinationText: comboLabel,
    breakdownTable,
    topLiftShare: topDriver?.liftShare || 0,
    sampleSize: topDriver?.sampleSize || 0,
    varianceScore,
  };
}

function buildInsightDiagnostics({ rows, metric, metricType, comparisonTable, driverInfo, deltaPercent, dimension }) {
  const anomalies = [];
  const concentrationRisk = [];
  const relationships = [];
  const topShare = comparisonTable?.[0]?.rawValue && comparisonTable.length
    ? (comparisonTable[0].rawValue / Math.max(1, comparisonTable.reduce((s, r) => s + (r.rawValue || 0), 0)))
    : 0;
  if (topShare > 0.6) concentrationRisk.push(`${comparisonTable[0].segment} accounts for ${percentFormatter.format(topShare)} of ${formatMetricLabel(metric)} in ${dimension}.`);
  if (Math.abs(deltaPercent || 0) < 10) anomalies.push("Performance gap is small; treat the insight as directional.");
  if (driverInfo?.varianceScore != null && driverInfo.varianceScore > 0.4) anomalies.push("Driver trend is volatile across observed periods.");
  const corrCandidates = (state.schema.numeric || []).filter((m) => m !== metric);
  let best = null;
  corrCandidates.forEach((candidate) => {
    const corr = computeCorrelation(rows, metric, candidate);
    if (corr === null) return;
    if (!best || Math.abs(corr) > Math.abs(best.corr)) best = { candidate, corr };
  });
  if (best) relationships.push(`${formatMetricLabel(best.candidate)} has the strongest relationship to ${formatMetricLabel(metric)} (r=${best.corr.toFixed(2)}).`);
  if (!relationships.length) relationships.push("No strong secondary metric relationships were detected.");
  if (!concentrationRisk.length) concentrationRisk.push("No severe concentration risk detected in the current breakdown.");
  if (!anomalies.length) anomalies.push("No major anomalies detected in the current slice.");
  return { anomalies, concentrationRisk, relationships };
}

function classifyNumericMetricColumn(colName, profile, rowCount) {
  const lower = String(colName || "").toLowerCase();
  const nonBlank = profile?.nonBlankCount || rowCount || 0;
  const uniqueRate = profile?.uniqueRate ?? 0;
  const isIdentifierLike = isIdLikeColumn(colName, profile, rowCount)
    || /\b(employee|account|member|user[_\s-]?id|number)\b/i.test(lower)
    || (nonBlank >= 10 && uniqueRate > 0.9);
  const isCurrency = /(revenue|sales|cost|spend|profit|price|amount|gmv|mrr|arr|pipeline)/i.test(lower);
  const isPercent = /(rate|percent|ctr|cvr|roas|ratio|retention|churn)/i.test(lower)
    || ((profile?.min ?? 0) >= 0 && (profile?.max ?? 2) <= 1.2 && !isIdentifierLike);
  const isDuration = /(time|duration|seconds|secs|minutes|mins|hours)/i.test(lower);
  const kind = isCurrency ? "currency" : isPercent ? "percent" : isDuration ? "duration" : "count";
  const kindPriority = isCurrency ? 1 : isPercent ? 0.85 : kind === "count" ? 0.7 : isDuration ? 0.55 : 0.4;
  return { isIdentifierLike, kind, kindPriority };
}

function formatDriverCombinationLabel(parts) {
  const clean = (parts || []).filter(Boolean);
  if (!clean.length) return "Unknown";
  if (clean.length === 1) return clean[0];
  if (clean.length === 2) return `${clean[0]} in ${clean[1]}`;
  return `${clean[0]} in ${clean[1]} via ${clean.slice(2).join(" / ")}`;
}

function computeUserUnifiedConfidence({ periods, sampleSize, deltaPercent, variance, hasDateField }) {
  let level = "high";
  const reasons = [];
  if (!hasDateField) {
    level = "medium";
    reasons.push("No historical time buckets available.");
  }
  if (periods < 3) {
    level = "low";
    reasons.push("Not enough time periods to confirm a stable trend.");
  }
  if (sampleSize < 8 && level !== "low") {
    level = "medium";
    reasons.push("Limited sample size in the top driver combination.");
  }
  if ((variance || 0) > 0.4 && level !== "low") {
    level = "medium";
    reasons.push("Trend shows volatility across periods.");
  }
  if ((deltaPercent || 0) < 10 && level !== "low") {
    level = "medium";
    reasons.push("Performance gap is relatively small.");
  }
  return {
    level,
    reasons: reasons.slice(0, 2),
    metrics: { periods, sampleSize, deltaPercent, variance: variance || 0 },
  };
}

function buildUserDynamicAction({ metricKey, topSegment, driverCombination, confidence, deltaStrength }) {
  const name = formatMetricLabel(metricKey);
  if (confidence?.level === "low") {
    return `Validate whether ${driverCombination} remains the primary driver with additional time periods before reallocating budget.`;
  }
  if (/retention|churn|rate/i.test(name)) {
    return `Identify what is unique about ${driverCombination} and apply the same lifecycle tactics to lift performance outside ${topSegment}.`;
  }
  if (/revenue|sales|arr|mrr|gmv/i.test(name)) {
    return `Rebalance investment toward ${driverCombination} while protecting efficiency, then test expansion into the next best segment.`;
  }
  if (/pipeline|opportunit/i.test(name)) {
    return `Inspect funnel stage conversion for ${driverCombination} and replicate the strongest stage improvements in lower performing segments.`;
  }
  if ((deltaStrength || 0) < 10) {
    return `Monitor performance before reallocating budget due to a small performance gap.`;
  }
  return `Investigate drivers behind ${driverCombination} and run a controlled test to validate impact.`;
}

function buildDecisionInsights(rows) {
  const insights = [];
  const metric = state.selections.primaryMetric;
  if (!metric) return insights;
  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  const segmentColumn = state.industryColumn || chooseCategoryColumn(rows, state.schema.profiles);

  const coverage = {
    rows: rows.length,
    dateRange: state.dateColumn && state.schema.profiles[state.dateColumn]?.dateMin
      ? `${formatDateInput(state.schema.profiles[state.dateColumn].dateMin)} → ${formatDateInput(state.schema.profiles[state.dateColumn].dateMax)}`
      : "No date column",
    dimensions: (state.schema.categoricals || []).slice(0, 3).join(", ") || "None",
  };

  const primaryValue = metricType.kind === "rate"
    ? (metricValues.reduce((sum, val) => sum + val, 0) / (metricValues.length || 1))
    : metricValues.reduce((sum, val) => sum + val, 0);

  const summaryBullets = [
    `Why: Primary metric is ${formatMetricValue(primaryValue, metricType)} across ${coverage.rows} rows.`,
    `Evidence: Date range ${coverage.dateRange}.`,
    `So what: Key dimensions available — ${coverage.dimensions}.`,
    `Next step: Validate if ${segmentColumn || "segments"} explain variation.`,
  ];
  insights.push({
    title: "Executive summary",
    summary: `${metric} coverage and baseline.`,
    bullets: summaryBullets,
  });

  if (segmentColumn) {
    const breakdown = aggregateByCategory(rows, segmentColumn, metric, 12, metricType);
    if (breakdown.labels.length >= 2) {
      const top = { label: breakdown.labels[0], value: breakdown.values[0], n: breakdown.counts[0] };
      const bottom = {
        label: breakdown.labels[breakdown.labels.length - 1],
        value: breakdown.values[breakdown.values.length - 1],
        n: breakdown.counts[breakdown.counts.length - 1],
      };
      const delta = top.value - bottom.value;
      const deltaLabel = metricType.kind === "rate" ? `${(delta * 100).toFixed(1)} pp` : formatMetricValue(delta, metricType);
      insights.push({
        title: "Segment lift",
        summary: `${top.label} vs ${bottom.label} on ${metric}.`,
        bullets: [
          `Why: ${top.label} is highest while ${bottom.label} is lowest.`,
          `Evidence: Top ${formatMetricValue(top.value, metricType)} (n=${top.n}); Bottom ${formatMetricValue(bottom.value, metricType)} (n=${bottom.n}).`,
          `So what: Absolute gap is ${deltaLabel}.`,
          `Next step: Compare drivers for ${top.label} vs ${bottom.label}.`,
        ],
      });
    }
  }

  if (state.dateColumn) {
    const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange, metricType);
    const window = series.values.length >= 14 ? 7 : series.values.length >= 6 ? 3 : 0;
    if (window) {
      const recent = series.values.slice(-window);
      const prior = series.values.slice(-window * 2, -window);
      if (prior.length) {
        const recentAvg = recent.reduce((sum, val) => sum + val, 0) / recent.length;
        const priorAvg = prior.reduce((sum, val) => sum + val, 0) / prior.length;
        const delta = recentAvg - priorAvg;
        const deltaLabel = metricType.kind === "rate" ? `${(delta * 100).toFixed(1)} pp` : formatMetricValue(delta, metricType);
        insights.push({
          title: "Period over period",
          summary: `${metric} moved ${deltaLabel} vs prior window.`,
          bullets: [
            `Why: Recent ${window} periods differ from prior ${window}.`,
            `Evidence: Recent ${formatMetricValue(recentAvg, metricType)} vs Prior ${formatMetricValue(priorAvg, metricType)}.`,
            `So what: Change = ${deltaLabel}.`,
            `Next step: Identify which segments shifted during this window.`,
          ],
        });
      }
    }
  }

  let driverInsight = null;
  const numericCandidates = (state.schema.numeric || []).filter((m) => m !== metric);
  let best = null;
  numericCandidates.forEach((candidate) => {
    const corr = computeCorrelation(rows, metric, candidate);
    if (corr === null) return;
    if (!best || Math.abs(corr) > Math.abs(best.corr)) best = { metric: candidate, corr };
  });
  if (best && rows.length >= 20) {
    driverInsight = {
      title: "Drivers and relationships",
      summary: `${best.metric} shows the strongest relationship to ${metric}.`,
      bullets: [
        `Why: Correlation suggests directional influence.`,
        `Evidence: Correlation r=${best.corr.toFixed(2)}.`,
        `So what: Changes in ${best.metric} align with ${metric}.`,
        `Next step: Validate causality with a controlled slice on ${best.metric}.`,
      ],
    };
  } else {
    driverInsight = {
      title: "Drivers and relationships",
      summary: "Insufficient data to validate numeric drivers.",
      bullets: [
        "Why: Fewer than 20 clean rows or no strong numeric pairs.",
        "Evidence: Driver analysis skipped.",
        "So what: Focus on segment comparisons first.",
        "Next step: Collect more rows or add numeric fields like Sessions or ActiveUsers.",
      ],
    };
  }
  insights.push(driverInsight);

  if (state.dateColumn) {
    const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange, metricType);
    if (series.values.length >= 8) {
      const mean = series.values.reduce((sum, val) => sum + val, 0) / series.values.length;
      const variance = series.values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / series.values.length;
      const std = Math.sqrt(variance);
      const anomalies = series.values
        .map((value, idx) => ({ value, label: series.labels[idx] }))
        .filter((point) => Math.abs(point.value - mean) > 2 * std);
      if (anomalies.length) {
        insights.push({
          title: "Anomalies",
          summary: `${anomalies.length} periods exceed 2σ from mean.`,
          bullets: [
            `Why: Values deviate beyond 2 standard deviations.`,
            `Evidence: Example ${anomalies[0].label} = ${formatMetricValue(anomalies[0].value, metricType)}.`,
            "So what: Spikes/drops may reflect campaign or product changes.",
            "Next step: Review logs or notes for those dates.",
          ],
        });
      }
    }
  }

  if (segmentColumn) {
    const breakdown = aggregateByCategory(rows, segmentColumn, metric, 8, metricType);
    if (breakdown.labels.length) {
      const total = breakdown.values.reduce((sum, val) => sum + val, 0);
      const share = total ? breakdown.values[0] / total : 0;
      if (share > 0.4) {
        insights.push({
          title: "Concentration risk",
          summary: `${breakdown.labels[0]} contributes ${formatMetric(share, "rate")}.`,
          bullets: [
            `Why: Top segment exceeds 40% share.`,
            `Evidence: Share = ${formatMetric(share, "rate")}.`,
            "So what: Performance is concentrated and brittle.",
            "Next step: Diversify growth across other segments.",
          ],
        });
      }
    }
  }

  const actions = [];
  const preferredDim = segmentColumn
    || (state.schema.categoricals || []).find((col) => /plan|region/i.test(col))
    || (state.schema.categoricals || [])[0]
    || "segment";
  actions.push(`Investigate ${preferredDim} segments where ${metric} is lowest and compare Sessions/ActiveUsers to spot friction.`);
  actions.push(`Run a cohort review for ${preferredDim} segments with low ${metric} and validate against ActiveUsers trends.`);
  const driverMetric = best?.metric || "Sessions";
  actions.push(`Quantify how changes in ${driverMetric} shift ${metric} and test a targeted uplift plan.`);

  insights.push({
    title: "Action items",
    summary: "Three next steps tied to observed segments/drivers.",
    bullets: actions.slice(0, 3).map((action) => `Next step: ${action}`),
  });

  return insights;
}

function renderMetricComparison(rows) {
  if (!rows.length) return;
  prepareMetricComparisonNarrativeLayout();
  const numericMetrics = Array.isArray(state.schema.numeric) ? state.schema.numeric : [];
  if (!numericMetrics.length) return;

  const primary = state.selections.primaryMetric || chooseTopNumericByVariance(state.schema.profiles, numericMetrics);
  state.selections.primaryMetric = primary;

  if (primaryMetricSelect) {
    primaryMetricSelect.innerHTML = "";
    (numericMetrics || []).forEach((metric) => {
      const option = document.createElement("option");
      option.value = metric;
      option.textContent = metric;
      primaryMetricSelect.appendChild(option);
    });
    primaryMetricSelect.value = primary;
    primaryMetricSelect.onchange = () => {
      state.selections.primaryMetric = primaryMetricSelect.value;
      applyFiltersAndRender();
    };
  }

  const availableCompare = numericMetrics.filter((metric) => metric !== primary);
  if (compareMetricSelect) {
    const selected = chooseTopMetricDriversByCorrelation(rows, primary, availableCompare, 3);
    state.selections.compareMetrics = Array.isArray(selected) ? selected : [];
    compareMetricSelect.innerHTML = "";
    (availableCompare || []).forEach((metric) => {
      const option = document.createElement("option");
      option.value = metric;
      option.textContent = metric;
      option.selected = selected.includes(metric);
      compareMetricSelect.appendChild(option);
    });
    compareMetricSelect.onchange = null;
  }

  const narrative = buildMetricComparisonNarrative(rows, primary, state.selections.compareMetrics);
  renderMetricComparisonNarrativeHeader(primary, narrative);
  renderComparisonTrend(rows, primary);
  renderDriversList(rows, primary, state.selections.compareMetrics, narrative);
  hideMetricComparisonSegmentView();
}

function renderComparisonTrend(rows, primary) {
  if (comparisonChartInstance) comparisonChartInstance.destroy();
  if (!comparisonTrendChart) return;
  const metricValues = rows.map((row) => parseNumber(row[primary])).filter((value) => value !== null);
  const metricType = inferMetricType(primary, metricValues);

  if (state.dateColumn) {
    const series = aggregateByDate(rows, state.dateColumn, primary, "day", state.dateRange);
    comparisonChartInstance = createLineChart("comparisonTrendChart", series.labels, series.values, metricType);
    if (comparisonNote) comparisonNote.textContent = `Primary metric trend by ${state.dateColumn}`;
  } else {
    const category = chooseCategoryColumn(rows, state.schema.profiles);
    if (category) {
      const series = aggregateByCategory(rows, category, primary, 8);
      comparisonChartInstance = createBarChart("comparisonTrendChart", series.labels, series.values, metricType);
      if (comparisonNote) comparisonNote.textContent = `Primary metric by ${category}`;
    } else {
      const labels = rows.map((_, index) => index + 1);
      const values = rows.map((row) => parseNumber(row[primary]) ?? 0);
      comparisonChartInstance = createLineChart("comparisonTrendChart", labels, values, metricType);
      if (comparisonNote) comparisonNote.textContent = "Primary metric by row order";
    }
  }
}

function renderDriversList(rows, primary, compareMetrics, narrative = null) {
  if (!driversList) return;
  driversList.innerHTML = "";
  const metrics = Array.isArray(compareMetrics) ? compareMetrics : [];
  if (!metrics.length) {
    driversList.innerHTML = "<li>No strong numeric drivers were detected.</li>";
    return;
  }

  const items = metrics.map((metric) => {
    const corr = computeCorrelation(rows, primary, metric);
    if (corr === null) {
      return { metric, corr: null, label: "Insufficient data" };
    }
    return { metric, corr };
  }).sort((a, b) => Math.abs(b.corr ?? 0) - Math.abs(a.corr ?? 0));

  const actionLine = narrative?.action
    ? `<li><strong>Action:</strong> ${escapeHtml(narrative.action)}</li>`
    : "";
  const explanationLines = items.slice(0, 3).map((item) => {
    if (item.corr === null) return `<li>${escapeHtml(formatMetricLabel(item.metric))}: insufficient data for interpretation.</li>`;
    const percent = Math.round(Math.abs(item.corr) * 100);
    const directionText = item.corr >= 0 ? "increases" : "decreases";
    return `<li>${escapeHtml(formatMetricLabel(item.metric))} explains ${percent}% of ${escapeHtml(formatMetricLabel(primary))} variation and typically ${directionText} with it.</li>`;
  }).join("");

  driversList.innerHTML = `
    ${narrative?.driverExplanation ? `<li>${escapeHtml(narrative.driverExplanation)}</li>` : ""}
    ${explanationLines}
    ${actionLine}
  `;
}

function renderSegmentView(rows, primary) {
  if (!segmentChartCanvas) return;
  if (segmentChartInstance) segmentChartInstance.destroy();
  const metricValues = rows.map((row) => parseNumber(row[primary])).filter((value) => value !== null);
  const metricType = inferMetricType(primary, metricValues);

  let dimension = state.industryColumn;
  if (state.filters.industry && state.filters.industry !== "All") {
    dimension = findAlternateCategory(state.schema.categoricals, state.industryColumn);
  }
  if (!dimension) {
    dimension = chooseCategoryColumn(rows, state.schema.profiles);
  }
  if (!dimension) return;

  const breakdown = aggregateByCategory(rows, dimension, primary, 8);
  segmentChartInstance = createBarChart("segmentChart", breakdown.labels, breakdown.values, metricType);
}

function prepareMetricComparisonNarrativeLayout() {
  const comparisonSectionEl = document.getElementById("comparisonSection");
  if (!comparisonSectionEl) return;
  const title = comparisonSectionEl.querySelector(".chart-header h3");
  const subtitle = comparisonSectionEl.querySelector(".chart-header p");
  if (title) {
    const metricLabel = formatMetricLabel(state.selections.primaryMetric || "");
    title.textContent = metricLabel ? `What drives ${metricLabel}` : "What drives this metric";
  }
  if (subtitle) {
    subtitle.textContent = "One interpreted view of the strongest drivers and what to do next.";
  }

  const compareGroup = compareMetricSelect?.closest(".control-group");
  if (compareGroup) compareGroup.classList.add("hidden");

  const segmentPanel = segmentChartCanvas?.closest(".comparison-card");
  if (segmentPanel) segmentPanel.classList.add("hidden");

  const driverPanel = driversList?.closest(".comparison-card");
  if (driverPanel) {
    const h4 = driverPanel.querySelector("h4");
    if (h4) h4.textContent = "Driver explanation";
  }
}

function renderMetricComparisonNarrativeHeader(primary, narrative) {
  const wrapper = document.querySelector("#comparisonSection .heatmap-wrapper");
  if (!wrapper) return;
  let narrativeBlock = document.getElementById("comparisonNarrativeHeader");
  if (!narrativeBlock) {
    narrativeBlock = document.createElement("div");
    narrativeBlock.id = "comparisonNarrativeHeader";
    narrativeBlock.className = "comparison-narrative";
    wrapper.parentNode.insertBefore(narrativeBlock, wrapper);
  }
  const primaryLabel = formatMetricLabel(primary);
  const headline = narrative?.headline || `This view highlights the strongest measurable drivers of ${primaryLabel}.`;
  const action = narrative?.action || `Validate the top driver for ${primaryLabel} before making changes.`;
  narrativeBlock.innerHTML = `
    <p class="comparison-narrative-headline">${escapeHtml(headline)}</p>
    <p class="comparison-narrative-action"><strong>Action:</strong> ${escapeHtml(action)}</p>
  `;
}

function chooseTopMetricDriversByCorrelation(rows, primary, candidates, limit = 3) {
  return [...(candidates || [])]
    .map((metric) => ({ metric, corr: computeCorrelation(rows, primary, metric) }))
    .filter((item) => item.corr !== null)
    .sort((a, b) => Math.abs(b.corr) - Math.abs(a.corr))
    .slice(0, limit)
    .map((item) => item.metric);
}

function buildMetricComparisonNarrative(rows, primary, compareMetrics) {
  const primaryLabel = formatMetricLabel(primary);
  const topDriver = (compareMetrics || [])[0];
  const topCorr = topDriver ? computeCorrelation(rows, primary, topDriver) : null;
  const topCorrPct = topCorr === null ? null : Math.round(Math.abs(topCorr) * 100);
  const directionWord = (topCorr ?? 0) >= 0 ? "rises" : "falls";
  const segmentDim = state.selectedDimension || chooseCategoryColumn(rows, state.schema.profiles);
  let topSegments = [];
  if (segmentDim) {
    const metricValues = rows.map((row) => parseNumber(row[primary])).filter((value) => value !== null);
    const metricType = inferMetricType(primary, metricValues);
    const breakdown = aggregateByCategory(rows, segmentDim, primary, 2, metricType);
    topSegments = (breakdown.labels || []).slice(0, 2);
  }
  const regionHint = topSegments.length ? `, especially in ${topSegments.join(" and ")}` : "";
  const headline = topDriver && topCorr !== null
    ? `${formatMetricLabel(topDriver)} ${directionWord} with ${primaryLabel}${regionHint}.`
    : `${primaryLabel} is shown in a single trend view${regionHint}.`;
  const driverExplanation = topDriver && topCorr !== null
    ? `${formatMetricLabel(topDriver)} explains ${topCorrPct}% of ${primaryLabel} variation${topCorr >= 0 ? "" : " (inverse relationship)"} based on correlation strength.`
    : `No strong numeric driver was detected for ${primaryLabel} in the current slice.`;
  const action = topDriver && topCorr !== null && Math.abs(topCorr) >= 0.6
    ? `Review cases where ${formatMetricLabel(topDriver)} changes faster than ${primaryLabel} to confirm whether the driver is causal.`
    : `Validate the top driver relationship with a filtered segment before acting on ${primaryLabel}.`;
  return { headline, driverExplanation, action };
}

function hideMetricComparisonSegmentView() {
  const segmentPanel = segmentChartCanvas?.closest(".comparison-card");
  if (segmentPanel) segmentPanel.classList.add("hidden");
  if (segmentChartInstance) {
    segmentChartInstance.destroy();
    segmentChartInstance = null;
  }
}

function findAlternateCategory(categories, industryColumn) {
  return categories.find((col) => col !== industryColumn) || null;
}

function chooseTopNumericByVariance(profile, numericColumns, limit = 1) {
  const sorted = [...(numericColumns || [])].sort((a, b) => (profile[b]?.std || 0) - (profile[a]?.std || 0));
  if (limit === 1) return sorted[0];
  return sorted.slice(0, limit);
}

function renderSuggestedTrends() {
  if (!suggestedTrends) return;
  suggestedTrends.innerHTML = "";
  if (!state.rawRows.length) {
    suggestedTrends.innerHTML = "<li>Load a sample to see suggested trends.</li>";
    return;
  }
  const suggestions = [];
  if (state.dateColumn && state.schema.numeric.length) {
    const metric = state.selections.primaryMetric || state.schema.numeric[0];
    suggestions.push(`${metric} trend over ${state.dateColumn}`);
    if (state.schema.categoricals.length) {
      suggestions.push(`${metric} by ${state.schema.categoricals[0]}`);
    }
  } else if (state.schema.numeric.length && state.schema.categoricals.length) {
    suggestions.push(`${state.selections.primaryMetric || state.schema.numeric[0]} by ${state.schema.categoricals[0]}`);
  }

  if (!suggestions.length) {
    suggestions.push("Add a date or categorical column to unlock trend suggestions.");
  }

  (suggestions || []).forEach((text) => {
    const li = document.createElement("li");
    li.textContent = text;
    suggestedTrends.appendChild(li);
  });
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
  const insights = buildDecisionInsights(rows);
  const content = [
    "Insights",
    `Generated: ${new Date().toLocaleString()}`,
    `Metric: ${state.selections.primaryMetric}`,
    `Dimension: ${state.selectedDimension || "—"}`,
    "",
    ...insights.map((insight, index) => `${index + 1}. ${insight.title} - ${insight.summary}`),
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
