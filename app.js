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
const includeDiagnosticsExport = document.getElementById("includeDiagnosticsExport");
const diagnosticsSection = document.getElementById("diagnosticsSection");
const exportStatus = document.getElementById("exportStatus");
const webhookStatus = document.getElementById("webhookStatus");
const stateSection = document.getElementById("state");
const errorSection = document.getElementById("error");
const warningsSection = document.getElementById("warnings");
const debugUploadStatus = document.getElementById("debugUploadStatus");
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
const exportCsvButton = document.getElementById("exportCsv");
const downloadInsightsButton = document.getElementById("downloadInsights");
const trendTitle = document.getElementById("trendTitle");
const trendSubtitle = document.getElementById("trendSubtitle");
const trendEmptyState = document.getElementById("trendEmptyState");
const barTitle = document.getElementById("barTitle");
const barSubtitle = document.getElementById("barSubtitle");
const insightsList = document.getElementById("insightsList");
const profileTable = document.getElementById("profileTable");
const profileSummaryLine = document.getElementById("profileSummaryLine");
const qualityBadge = document.getElementById("qualityBadge");
const buildStamp = document.getElementById("buildStamp");
const dataSummaryPanel = document.getElementById("dataSummaryPanel");
const previewCompactToggle = document.getElementById("previewCompactToggle");
const diagCorrTitle = document.getElementById("diagCorrTitle");
const diagCorrSubtitle = document.getElementById("diagCorrSubtitle");
const diagCorrExplain = document.getElementById("diagCorrExplain");
const diagCorrTakeaway = document.getElementById("diagCorrTakeaway");
const diagCorrCaution = document.getElementById("diagCorrCaution");
const diagCorrXSelect = document.getElementById("diagCorrX");
const diagCorrYSelect = document.getElementById("diagCorrY");
const diagCorrSummary = document.getElementById("diagCorrSummary");
const diagCorrChartCanvas = document.getElementById("diagCorrChart");
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
const modalUploadInput = document.getElementById("modalUploadInput");
const modalPdfInput = document.getElementById("modalPdfInput");
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
const sampleKpiGrid = document.getElementById("sampleKpiGrid");
const sampleTrendMetric = document.getElementById("sampleTrendMetric");
const sampleTrendBreakdown = document.getElementById("sampleTrendBreakdown");
const sampleTrendChart = document.getElementById("sampleTrendChart");
const sampleInsightsRoot = document.getElementById("sampleInsightsRoot");
const sampleDimensionSelect = document.getElementById("sampleDimensionSelect");
const sampleBreakdownChart = document.getElementById("sampleBreakdownChart");
const suggestedTrends = document.getElementById("suggestedTrends");
const chatPanel = document.getElementById("chatPanel");
const chatInput = document.getElementById("chatInput");
const chatSendButton = document.getElementById("chatSend");
const chatStatus = document.getElementById("chatStatus");
const chatPreview = document.getElementById("chatPreview");
const chatApplyButton = document.getElementById("chatApply");
const chatCancelButton = document.getElementById("chatCancel");
const chatUndoButton = document.getElementById("chatUndo");
const chatRedoButton = document.getElementById("chatRedo");
const chatActivity = document.getElementById("chatActivity");
const evidenceDrawer = document.getElementById("evidenceDrawer");
const evidenceCloseButton = document.getElementById("evidenceClose");
const evidenceContent = document.getElementById("evidenceContent");
const evidenceScrim = document.getElementById("evidenceScrim");
const headerTabsContainer = document.querySelector(".header-tabs");

let dashboardStateModule = null;
let chartActionsModule = null;
let schemaGuardsModule = null;
let detectDateFieldModule = null;
let chatPreviewState = null;
let pendingToolCalls = [];
let pendingAssistantMessage = "";

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
let diagnosticsCorrelationChartInstance = null;
let heatmapChartInstance = null;
let diagnosticsCorrelationResizeObserver = null;
let currentTableRows = [];
let currentSort = { key: null, direction: "asc" };
let currentPreviewColumns = [];
let currentPreviewBaseRows = [];
let googleAccessToken = null;
let googleTokenClient = null;
let logoWarningShown = false;
let renderKpiCallCount = 0;
let signupSubmitting = false;
let signupMode = "signup";
let currentUserProfile = null;
let diagnosticsOpenBeforePrint = null;
const signupFormState = {
  fullName: "",
  email: "",
  company: "",
  useCase: "",
  useCaseDetail: "",
};
const USERS_KEY = "shay_users";
const SESSION_KEY = "shay_session";

const DEMO_CSV_URL = new URL("./data-sample.csv", document.baseURI).toString();
const LOCAL_PATH_ERROR_MESSAGE = "Local file paths are not accessible in the browser. Please upload the file or host it in the repo.";
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

const MAX_HEATMAP_ROWS = 8;
const MAX_HEATMAP_COLS = 6;
const compactNumberFormatter = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 });
const compactCurrencyFormatter = new Intl.NumberFormat("en-US", {
  notation: "compact",
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 1,
});

const state = {
  dataset: null,
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
  selections: {
    primaryMetric: null,
    compareMetrics: [],
  },
  selectedMetric: null,
  selectedDimension: null,
  chartType: "line",
  aggregation: "sum",
  sort: null,
  chartState: null,
  domain: "general",
  domainAuto: true,
  inferredDomain: "general",
  chosenXAxisType: "category",
  dateColumn: null,
  dateFieldConfidence: 0,
  topN: 8,
  diagnosticsCorrelation: {
    xMetric: null,
    yMetric: null,
  },
  previewCompact: true,
  sampleTabAutoloaded: false,
  uiMode: "empty", // empty | loading | ready | error
  parseStatus: "idle",
  error: null,
};

const landingDemoState = {
  rows: [],
  columns: [],
  profiles: {},
  numericColumns: [],
  categoricalColumns: [],
  dateColumn: null,
  selectedMetric: null,
  selectedBreakdown: "overall",
  selectedDimension: null,
  breakdownMap: {
    plan: null,
    channel: null,
  },
  kpiCards: [],
};

const appStore = (() => {
  let store = {
    dataset: null,
    source: null,
    parseStatus: "idle", // idle | parsing | ready | error
    error: null,
    view: "upload",
    selections: {
      primaryMetric: null,
      breakdown: null,
      topN: 8,
    },
  };
  const listeners = new Set();
  return {
    getState() {
      return store;
    },
    setState(partial) {
      store = {
        ...store,
        ...partial,
        selections: {
          ...store.selections,
          ...(partial?.selections || {}),
        },
      };
      console.log("STORE update", {
        parseStatus: store.parseStatus,
        source: store.source,
        rows: store.dataset?.rows?.length || 0,
        columns: store.dataset?.columns?.length || 0,
        error: store.error || null,
        view: store.view,
      });
      listeners.forEach((listener) => {
        try {
          listener(store);
        } catch (error) {
          console.error("store listener error", error);
        }
      });
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
})();

function isAnalysisReady(storeState = appStore.getState()) {
  return storeState.parseStatus === "ready" && Boolean(storeState.dataset?.rows?.length);
}

window.__appStore = appStore;

document.addEventListener("DOMContentLoaded", init);
appStore.subscribe(() => {
  renderDebugUploadStatus();
  renderPdfStatusFromStore();
});

function renderDebugUploadStatus() {
  if (!debugUploadStatus) return;
  const storeState = appStore.getState();
  const dataset = storeState.dataset;
  const rawSourceLabel = typeof storeState.source === "string"
    ? storeState.source
    : (storeState.source?.type || "—");
  const sourceLabel = rawSourceLabel === "sheet" ? "sheets" : rawSourceLabel;
  debugUploadStatus.textContent = `Status: ${storeState.parseStatus} · Rows: ${dataset?.rows?.length || 0} · Columns: ${dataset?.columns?.length || 0} · Source: ${sourceLabel || "—"}`;
}

function setParseStatus(nextStatus, error = null) {
  state.parseStatus = nextStatus;
  state.error = error || null;
  appStore.setState({ parseStatus: nextStatus, error: error || null });
  renderDebugUploadStatus();
}

function renderPdfStatusFromStore() {
  if (!pdfStatus) return;
  const storeState = appStore.getState();
  const sourceType = typeof storeState.source === "string" ? storeState.source : storeState.source?.type;
  if (sourceType !== "pdf") return;
  if (storeState.parseStatus === "parsing") {
    pdfStatus.textContent = "Extracting tables from PDF...";
    return;
  }
  if (storeState.parseStatus === "ready") {
    pdfStatus.textContent = "Tables extracted. Building dashboard...";
    return;
  }
  if (storeState.parseStatus === "error") {
    pdfStatus.textContent = storeState.error || "PDF ingest failed.";
    return;
  }
  if (storeState.parseStatus === "idle") {
    pdfStatus.textContent = "";
  }
}

function init() {
  console.log("init started");
  if (buildStamp) {
    buildStamp.textContent = `Build: ${new Date().toISOString()}`;
  }

  applyBrandAssets();
  ensureUnifiedUserDashboardLayout();
  ensureDashboardModules();

  window.addEventListener("hashchange", handleRoute);
  window.addEventListener("beforeprint", () => {
    diagnosticsOpenBeforePrint = applyDiagnosticsExportState();
  });
  window.addEventListener("afterprint", () => {
    restoreDiagnosticsExportState(diagnosticsOpenBeforePrint);
    diagnosticsOpenBeforePrint = null;
  });

  hydrateSession();

  const isDev = location.hostname === "localhost" || location.hostname === "127.0.0.1";
  if (loadFixtureButton) {
    loadFixtureButton.classList.toggle("hidden", !isDev);
    loadFixtureButton.addEventListener("click", () => loadFixtureCsv());
  }

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.nav) return;
      switchTab(button.dataset.tab);
    });
  });
  if (headerTabsContainer) {
    headerTabsContainer.addEventListener("click", handleHeaderNavClick);
  }

  if (uploadTrigger && uploadInput) {
    uploadTrigger.addEventListener("click", () => uploadInput.click());
  }
  if (modalUploadTrigger && modalUploadInput) {
    modalUploadTrigger.addEventListener("click", () => modalUploadInput.click());
  }
  if (uploadCsvQuick && uploadInput) {
    uploadCsvQuick.addEventListener("click", () => uploadInput.click());
  }
  if (modalCsvQuick && modalUploadInput) {
    modalCsvQuick.addEventListener("click", () => modalUploadInput.click());
  }
  if (uploadPdfQuick && pdfInput) {
    uploadPdfQuick.addEventListener("click", () => pdfInput.click());
  }
  if (modalPdfQuick && (modalPdfInput || pdfInput)) {
    modalPdfQuick.addEventListener("click", () => (modalPdfInput || pdfInput).click());
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
  if (modalUploadInput) {
    modalUploadInput.addEventListener("change", (event) => handleFileSelection(event.target.files[0]));
  }
  if (modalPdfInput) {
    modalPdfInput.addEventListener("change", () => {
      const file = modalPdfInput.files?.[0];
      if (!file) return;
      console.log("Modal PDF input change", file.name, file.size);
      appStore.setState({
        dataset: null,
        source: { type: "pdf", name: file.name },
        parseStatus: "parsing",
        error: null,
        view: "upload",
      });
      setParseStatus("parsing");
      setUiMode("loading");
      renderPdfStatusFromStore();
      handlePdfUpload(file);
      modalPdfInput.value = "";
    });
  }

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
    syncChartStateDirect({ selectedMetric: metricSelect.value });
    applyFiltersAndRender();
  });

  dimensionSelect.addEventListener("change", () => {
    state.selectedDimension = dimensionSelect.value;
    syncChartStateDirect({ selectedDimension: dimensionSelect.value });
    applyFiltersAndRender();
  });

  topNSelect.addEventListener("change", () => {
    state.topN = Number(topNSelect.value);
    syncChartStateDirect({ topN: state.topN });
    applyFiltersAndRender();
  });

  if (previewCompactToggle) {
    previewCompactToggle.checked = state.previewCompact !== false;
    previewCompactToggle.addEventListener("change", () => {
      state.previewCompact = Boolean(previewCompactToggle.checked);
      renderTable(state.filteredRows || [], state.schema.columns || []);
    });
  }

  if (diagCorrXSelect) {
    diagCorrXSelect.addEventListener("change", () => {
      state.diagnosticsCorrelation.xMetric = diagCorrXSelect.value || null;
      renderDiagnosticsCorrelation(state.filteredRows || []);
    });
  }
  if (diagCorrYSelect) {
    diagCorrYSelect.addEventListener("change", () => {
      state.diagnosticsCorrelation.yMetric = diagCorrYSelect.value || null;
      renderDiagnosticsCorrelation(state.filteredRows || []);
    });
  }
  if (diagnosticsSection) {
    diagnosticsSection.addEventListener("toggle", () => {
      if (!diagnosticsSection.open) return;
      forceDiagnosticsCorrelationResize(50);
    });
  }
  window.addEventListener("resize", () => {
    scheduleDiagnosticsCorrelationResize();
  });

  domainSelect.addEventListener("change", () => {
    state.domainAuto = domainSelect.value === "auto";
    state.domain = state.domainAuto ? state.inferredDomain : domainSelect.value;
    applyFiltersAndRender();
  });

  exportCsvButton.addEventListener("click", () => exportCsvData());
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
  if (chatSendButton) chatSendButton.addEventListener("click", handleChatSend);
  if (chatApplyButton) chatApplyButton.addEventListener("click", applyChatPreview);
  if (chatCancelButton) chatCancelButton.addEventListener("click", clearChatPreview);
  if (chatUndoButton) chatUndoButton.addEventListener("click", () => handleUndoRedo("undo"));
  if (chatRedoButton) chatRedoButton.addEventListener("click", () => handleUndoRedo("redo"));
  if (evidenceCloseButton) evidenceCloseButton.addEventListener("click", closeEvidenceDrawer);
  if (evidenceScrim) evidenceScrim.addEventListener("click", closeEvidenceDrawer);
  if (pdfInput) {
    pdfInput.addEventListener("change", () => {
      const file = pdfInput.files?.[0];
      if (!file) return;
      console.log("PDF input change", file.name, file.size);
      appStore.setState({
        dataset: null,
        source: { type: "pdf", name: file.name },
        parseStatus: "parsing",
        error: null,
        view: "upload",
      });
      setParseStatus("parsing");
      setUiMode("loading");
      renderPdfStatusFromStore();
      if (pdfMeta) {
        const sizeKb = Math.round(file.size / 1024);
        pdfMeta.textContent = `${file.name} · ${sizeKb} KB`;
      }
      handlePdfUpload(file);
    });
  }

  updateSignInButton();
  wireNoLocalPathValidation();
  renderSampleGallery();
  renderLandingDemoPreview();
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
  setParseStatus("idle");
  renderPdfStatusFromStore();
  updateAnalysisHeaderState(false);
  syncUploadAnalysisState();
  console.log("handlers wired");
}

function ensureUnifiedUserDashboardLayout() {
  if (!dashboard) return;
  if (dashboard.querySelector("#unifiedDashboardShell")) return;

  const charts = dashboard.querySelector(".charts");
  const insightsCard = dashboard.querySelector(".insights-card");
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
  const duplicates = sectionHeader.querySelectorAll("#dashboardActionsToolbar");
  if (duplicates.length > 1) {
    duplicates.forEach((node, index) => {
      if (index > 0) node.remove();
    });
  }

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
      if (hasLoadedDataset()) {
        console.log("Start new upload clicked");
        resetStateForNewDataset();
        switchTab("upload");
      }
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
    uploadToggle.textContent = hasDataset ? "Start new upload" : "Upload data";
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
  if (!uploaderModal) {
    console.warn("Uploader modal not found; falling back to file picker");
    uploadInput?.click();
    return;
  }
  uploaderModal.classList.remove("hidden");
  uploaderModal.setAttribute("aria-hidden", "false");
  console.log("Uploader modal open", !uploaderModal.classList.contains("hidden"));
}

function closeUploaderModal() {
  if (!uploaderModal) return;
  uploaderModal.classList.add("hidden");
  uploaderModal.setAttribute("aria-hidden", "true");
}

function hasLoadedDataset() {
  const storeState = appStore.getState();
  const dataset = storeState.dataset;
  return Boolean((dataset?.rows && dataset.rows.length > 0) || (state.rawRows && state.rawRows.length > 0));
}

function hasUserDataLoaded() {
  const storeState = appStore.getState();
  if (!isAnalysisReady(storeState)) return false;
  const sourceType = typeof storeState.source === "string"
    ? storeState.source
    : (storeState.source?.type || storeState.dataset?.meta?.sourceType || null);
  return sourceType !== "demo";
}

function setUiMode(mode) {
  state.uiMode = mode;
  if (mode === "ready") closeUploaderModal();
  appStore.setState({ view: mode === "ready" ? "analysis" : "upload" });
  syncUploadAnalysisState();
  renderDebugUploadStatus();
}

function onDatasetReady(sourceMeta = {}) {
  setUiMode("ready");
  setParseStatus("ready");
  updateAnalysisHeaderState(true);
  tabPanels.forEach((panel) => panel.classList.add("hidden"));
  if (dashboard) dashboard.classList.remove("hidden");
  try {
    dashboard?.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (_) {
    dashboard?.scrollIntoView();
  }
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

async function ensureDashboardModules() {
  if (dashboardStateModule && chartActionsModule && schemaGuardsModule && detectDateFieldModule) return;
  try {
    const [stateMod, actionsMod, guardsMod, dateMod] = await Promise.all([
      import("./src/state/dashboardState.js"),
      import("./src/actions/chartActions.js"),
      import("./src/validation/schemaGuards.js"),
      import("./src/profiling/detectDateField.js"),
    ]);
    dashboardStateModule = stateMod;
    chartActionsModule = actionsMod;
    schemaGuardsModule = guardsMod;
    detectDateFieldModule = dateMod;
  } catch (error) {
    console.warn("Dashboard modules failed to load, chat features disabled.", error);
  }
}

function buildDatasetSchema() {
  const columns = (state.schema.columns || []).map((name) => ({
    name,
    type: state.schema.profiles?.[name]?.type || "string",
  }));
  return {
    columns,
    dateField: state.dateColumn || null,
  };
}

function buildChartStateSnapshot() {
  return {
    dataset: state.dataset || null,
    selectedMetric: state.selections.primaryMetric || null,
    selectedDimension: state.selectedDimension || null,
    chartType: state.chartType || "bar",
    sort: state.sort || null,
    topN: state.topN || 8,
    aggregation: state.aggregation || "sum",
    timeField: state.dateColumn || null,
  };
}

function ensureChartState() {
  if (!dashboardStateModule?.createDashboardState) return;
  if (!state.chartState) {
    state.chartState = dashboardStateModule.createDashboardState(buildChartStateSnapshot());
  }
}

function syncChartStateDirect(patch = {}) {
  ensureChartState();
  if (!state.chartState) return;
  state.chartState = {
    ...state.chartState,
    ...patch,
  };
}

function applyChartState(nextState, options = {}) {
  if (!nextState) return;
  state.chartState = nextState;
  state.selections.primaryMetric = nextState.selectedMetric || state.selections.primaryMetric;
  state.selectedMetric = state.selections.primaryMetric;
  state.selectedDimension = nextState.selectedDimension || state.selectedDimension;
  state.chartType = nextState.chartType || state.chartType;
  state.aggregation = nextState.aggregation || state.aggregation;
  state.sort = nextState.sort || null;
  state.topN = Number.isFinite(nextState.topN) ? nextState.topN : state.topN;
  state.dateColumn = nextState.timeField || state.dateColumn;

  if (metricSelect && state.selections.primaryMetric) {
    metricSelect.value = state.selections.primaryMetric;
  }
  if (dimensionSelect) {
    dimensionSelect.value = state.selectedDimension || "";
  }
  if (topNSelect) {
    topNSelect.value = String(state.topN || 8);
  }

  applyFiltersAndRender();
  if (options?.logMessage) logChatActivity(options.logMessage);
}

function setChatStatus(message, tone = "info") {
  if (!chatStatus) return;
  chatStatus.textContent = message || "";
  chatStatus.classList.toggle("error", tone === "error");
}

function logChatActivity(message) {
  if (!chatActivity) return;
  const item = document.createElement("div");
  item.className = "chat-activity-item";
  item.textContent = message;
  chatActivity.prepend(item);
}

function clearChatPreview() {
  chatPreviewState = null;
  pendingToolCalls = [];
  pendingAssistantMessage = "";
  if (chatPreview) {
    chatPreview.classList.add("hidden");
    chatPreview.innerHTML = "";
  }
  if (chatApplyButton) chatApplyButton.disabled = true;
  if (chatCancelButton) chatCancelButton.disabled = true;
}

function renderChatPreview(diff, assistantMessage) {
  if (!chatPreview) return;
  const diffLines = diff.map((item) => `<div><strong>${escapeHtml(item.label)}:</strong> ${escapeHtml(item.before)} → ${escapeHtml(item.after)}</div>`).join("");
  chatPreview.innerHTML = `
    <div class="chat-preview-message">${escapeHtml(assistantMessage || "Proposed changes")}</div>
    <div class="chat-preview-diff">${diffLines || "<div>No changes detected.</div>"}</div>
  `;
  chatPreview.classList.remove("hidden");
  if (chatApplyButton) chatApplyButton.disabled = false;
  if (chatCancelButton) chatCancelButton.disabled = false;
}

function buildPreviewDiff(currentState, nextState) {
  const fields = [
    { key: "selectedMetric", label: "Metric" },
    { key: "selectedDimension", label: "Dimension" },
    { key: "chartType", label: "Chart type" },
    { key: "aggregation", label: "Aggregation" },
    { key: "topN", label: "Top N" },
    { key: "sort", label: "Sort" },
    { key: "timeField", label: "Time field" },
  ];
  const diffs = [];
  fields.forEach((field) => {
    const before = currentState?.[field.key];
    const after = nextState?.[field.key];
    if (JSON.stringify(before) !== JSON.stringify(after)) {
      diffs.push({
        label: field.label,
        before: before == null ? "—" : String(before),
        after: after == null ? "—" : String(after),
      });
    }
  });
  return diffs;
}

function applyToolCall(stateSnapshot, toolCall) {
  const handlers = {
    setMetric: (s, args) => chartActionsModule.setMetric(s, args.field),
    setDimension: (s, args) => chartActionsModule.setDimension(s, args.field),
    setChartType: (s, args) => chartActionsModule.setChartType(s, args.type),
    setAggregation: (s, args) => chartActionsModule.setAggregation(s, args.agg),
    setTopN: (s, args) => chartActionsModule.setTopN(s, args.n),
    setSort: (s, args) => chartActionsModule.setSort(s, args.field, args.direction),
    undo: (s) => chartActionsModule.undo(s),
    redo: (s) => chartActionsModule.redo(s),
  };
  const handler = handlers[toolCall?.name];
  if (!handler) return null;
  return handler(stateSnapshot, toolCall.args || {});
}

function previewToolCalls(toolCalls, assistantMessage) {
  if (!schemaGuardsModule || !chartActionsModule) {
    setChatStatus("Chat tools are still loading. Try again in a moment.", "error");
    return;
  }
  ensureChartState();
  const schema = buildDatasetSchema();
  let nextState = state.chartState;
  for (const toolCall of toolCalls) {
    const validation = schemaGuardsModule.validateToolCall(schema, nextState, toolCall);
    if (!validation.ok) {
      setChatStatus(validation.message || "Invalid request.", "error");
      return;
    }
    const updated = applyToolCall(nextState, toolCall);
    if (!updated) {
      setChatStatus(`Unknown action: ${toolCall?.name}`, "error");
      return;
    }
    nextState = updated;
  }
  chatPreviewState = nextState;
  pendingToolCalls = toolCalls;
  pendingAssistantMessage = assistantMessage || "";
  const diff = buildPreviewDiff(state.chartState, nextState);
  renderChatPreview(diff, assistantMessage);
}

async function handleChatSend() {
  if (!chatInput) return;
  const message = chatInput.value.trim();
  if (!message) return;
  if (!state.dataset || !state.rawRows?.length) {
    setChatStatus("Load a dataset before using chat.", "error");
    return;
  }
  clearChatPreview();
  setChatStatus("Thinking...");
  await ensureDashboardModules();
  ensureChartState();
  const payload = {
    message,
    datasetSchema: buildDatasetSchema(),
    currentDashboardState: state.chartState,
  };
  try {
    const chatApiUrl = new URL("./api/chat", document.baseURI).toString();
    const response = await fetch(chatApiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const text = await response.text();
    let data = null;
    try {
      data = JSON.parse(text);
    } catch (error) {
      console.error("Chat response not JSON:", text.slice(0, 200));
      setChatStatus("Backend did not return JSON. Check proxy and backend status.", "error");
      return;
    }
    if (!response.ok || data?.ok === false) {
      setChatStatus(data?.error || "Chat request failed.", "error");
      return;
    }
    if (!Array.isArray(data.toolCalls)) {
      setChatStatus("Chat response missing tool calls.", "error");
      return;
    }
    setChatStatus("");
    previewToolCalls(data.toolCalls, data.assistantMessage);
  } catch (error) {
    setChatStatus(error?.message || "Chat request failed.", "error");
  }
}

function applyChatPreview() {
  if (!chatPreviewState) return;
  applyChartState(chatPreviewState, {
    logMessage: pendingAssistantMessage || "Applied chat edits.",
  });
  clearChatPreview();
}

function handleUndoRedo(kind) {
  if (!chartActionsModule) return;
  ensureChartState();
  const next = kind === "redo"
    ? chartActionsModule.redo(state.chartState)
    : chartActionsModule.undo(state.chartState);
  applyChartState(next, { logMessage: kind === "redo" ? "Redo applied." : "Undo applied." });
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
  const logoSrc = new URL("./Brand/logo.png", document.baseURI).toString();
  const logo192 = new URL("./Brand/logo-192.png", document.baseURI).toString();
  const logo512 = new URL("./Brand/logo-512.png", document.baseURI).toString();

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
  const correlationId = createIngestCorrelationId("csv");
  console.log(`[INGEST ${correlationId}] UPLOAD file selected`, file.name, file.size);
  console.log("file selected", file.name);
  clearMessages();
  appStore.setState({
    dataset: null,
    source: { type: "csv", name: file.name },
    parseStatus: "parsing",
    error: null,
    view: "upload",
  });
  setUiMode("loading");
  setParseStatus("parsing");
  setStatus("Reading file");
  const reader = new FileReader();
  reader.onload = () => {
    const text = reader.result;
    parseCsvText(text, {
      sourceType: "csv",
      name: file.name,
      correlationId,
    });
  };
  reader.onerror = () => {
    showError("Unable to read the file.", "Details: FileReader failed");
  };
  reader.readAsText(file);
}

function parseCsvText(text, sourceMeta = {}) {
  const correlationId = sourceMeta?.correlationId;
  console.log(correlationId ? `[INGEST ${correlationId}] PARSE start` : "PARSE start");
  setUiMode("loading");
  setParseStatus("parsing");
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
      console.log(correlationId ? `[INGEST ${correlationId}] PARSE complete` : "parse complete");
      console.log("PARSE raw rows", results.data?.length || 0);
      ingestDataset({
        source: sourceMeta?.sourceType || "csv",
        rows: results.data,
        meta: sourceMeta,
      });
    },
  });
}

function normalizePdfHeaderName(value, index) {
  let cleaned = String(value ?? "").trim().toLowerCase();
  if (!cleaned) cleaned = `column_${index + 1}`;
  cleaned = cleaned.replace(/[^a-z0-9]+/g, "_");
  cleaned = cleaned.replace(/^_+|_+$/g, "");
  cleaned = cleaned.replace(/_+/g, "_");
  if (!cleaned) cleaned = `column_${index + 1}`;
  return cleaned;
}

function inferPdfHeaders(rawRows) {
  const rows = Array.isArray(rawRows) ? rawRows : [];
  if (!rows.length) return { headers: [], headerIndex: 0 };
  const sample = rows.slice(0, 5);
  let bestIndex = 0;
  let bestScore = -Infinity;

  sample.forEach((row, index) => {
    const cells = Array.isArray(row) ? row : [];
    const nonEmpty = cells.filter((cell) => String(cell ?? "").trim() !== "");
    if (!cells.length || !nonEmpty.length) return;
    const numericCount = nonEmpty.filter((cell) => parseNumber(cell) !== null).length;
    const textDensity = nonEmpty.length / cells.length;
    const numericRatio = numericCount / nonEmpty.length;
    const score = (textDensity * 2) - (numericRatio * 3);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  const headers = (sample[bestIndex] || []).map((cell, index) => normalizePdfHeaderName(cell, index));
  return { headers, headerIndex: bestIndex };
}

function loadSampleGallery() {
  renderSampleGallery();
  switchTab("samples");
}

function handleHeaderNavClick(event) {
  const navButton = event.target?.closest?.("[data-nav]");
  if (!navButton || (headerTabsContainer && !headerTabsContainer.contains(navButton))) return;
  const navKey = String(navButton.dataset.nav || "").trim().toLowerCase();
  if (!navKey) return;
  console.log(`[HeaderNav] clicked: ${navKey}`);

  if (navKey === "sheets") {
    const panel = document.getElementById("sheetsPanel");
    if (!panel) {
      console.error("[HeaderNav] Missing #sheetsPanel. Cannot route to Google Sheets.");
      return;
    }
    switchTab("sheets");
    return;
  }

  if (navKey === "api") {
    const panel = document.getElementById("apiPanel");
    if (!panel) {
      console.error("[HeaderNav] Missing #apiPanel. Cannot route to API connect.");
      return;
    }
    switchTab("api");
    return;
  }

  if (navKey === "pdf") {
    if (!pdfInput) {
      console.error("[HeaderNav] Missing #pdfInput. Cannot trigger PDF upload.");
      return;
    }
    const panel = document.getElementById("pdfPanel");
    if (!panel) {
      console.error("[HeaderNav] Missing #pdfPanel. Cannot route to PDF panel.");
      return;
    }
    switchTab("pdf");
    pdfInput.click();
    return;
  }

  if (navKey === "samples") {
    const panel = document.getElementById("samplesPanel");
    if (!panel) {
      console.error("[HeaderNav] Missing #samplesPanel. Cannot route to Samples.");
      return;
    }
    renderSampleGallery();
    switchTab("samples");
    loadFixtureCsv();
    return;
  }

  console.error(`[HeaderNav] Unknown nav target: ${navKey}`);
}

function switchTab(tabName) {
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
    ["export-csv", exportCsvData],
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
  document.body.classList.add("exportMode");
  const previousDiagnosticsOpen = applyDiagnosticsExportState();
  try {
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
  } finally {
    restoreDiagnosticsExportState(previousDiagnosticsOpen);
    document.body.classList.remove("exportMode");
  }
}

function applyDiagnosticsExportState() {
  if (!diagnosticsSection) return null;
  const previous = diagnosticsSection.open;
  const includeDiagnostics = Boolean(includeDiagnosticsExport?.checked);
  diagnosticsSection.open = includeDiagnostics;
  return previous;
}

function restoreDiagnosticsExportState(previousOpen) {
  if (!diagnosticsSection) return;
  if (typeof previousOpen !== "boolean") return;
  diagnosticsSection.open = previousOpen;
}

async function handleApiFetch() {
  const base = apiBaseUrl?.value?.trim();
  const endpoint = apiEndpoint?.value?.trim();
  if (!base || !endpoint) {
    apiStatus.textContent = "Please enter Base URL and Endpoint to fetch data.";
    return;
  }
  if (!validateRemoteInput(base, "api") || !validateRemoteInput(endpoint, "api")) return;
  console.log("API selected", base, endpoint);
  appStore.setState({ source: { type: "api", name: `${base}/${endpoint}` }, parseStatus: "parsing", error: null });
  console.log("API parseStatus -> parsing");
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
  setUiMode("loading");
  setParseStatus("parsing");
  try {
    const response = await fetch(url, { headers });
    if (!response.ok) {
      apiStatus.textContent = `Request failed: ${response.status}`;
      appStore.setState({ dataset: null, source: { type: "api", name: `${base}/${endpoint}` }, parseStatus: "error", error: `Request failed: ${response.status}` });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError("API request failed.", `Details: ${response.status}`);
      return;
    }
    const json = await response.json();
    if (!Array.isArray(json) || !json.length || typeof json[0] !== "object") {
      apiStatus.textContent = "Expected a JSON array of objects.";
      appStore.setState({ dataset: null, source: { type: "api", name: `${base}/${endpoint}` }, parseStatus: "error", error: "Expected a JSON array of objects." });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError("API payload must be a JSON array of objects.");
      return;
    }
    console.log("API raw rows", json.length);
    apiStatus.textContent = "API data loaded";
    ingestJsonRows(json, { sourceType: "api", name: `${base}${endpoint}` });
    console.log("STORE setState parseStatus", appStore.getState().parseStatus);
  } catch (error) {
    apiStatus.textContent = "Network error while fetching API data.";
    appStore.setState({ dataset: null, source: { type: "api", name: `${base}/${endpoint}` }, parseStatus: "error", error: error.message || "Network error" });
    console.log("STORE setState parseStatus", appStore.getState().parseStatus);
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
  ingestDataset({
    source: sourceMeta?.sourceType || "api",
    rows,
    columns,
    meta: sourceMeta,
  });
}

function handleExportToSheets() {
  const user = currentUserProfile?.email || "anonymous";
  const filters = {};
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

function wireNoLocalPathValidation() {
  const validatePastedPath = (event) => {
    const target = event?.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) return;
    if (target.type === "file") return;
    const pasted = event?.clipboardData?.getData("text") || "";
    if (!isLikelyLocalWindowsPath(pasted)) return;
    target.setCustomValidity(LOCAL_PATH_ERROR_MESSAGE);
    if (typeof target.reportValidity === "function") target.reportValidity();
    showLocalPathMessageForInput(target);
  };
  document.addEventListener("paste", validatePastedPath);
  document.addEventListener("input", (event) => {
    const target = event?.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) return;
    if (target.validationMessage === LOCAL_PATH_ERROR_MESSAGE && !isLikelyLocalWindowsPath(target.value)) {
      target.setCustomValidity("");
    }
  });
}

function isLikelyLocalWindowsPath(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  return /^[a-zA-Z]:[\\/]/.test(text) || /^\\\\/.test(text) || /^file:\/\/\/[a-zA-Z]:\//i.test(text);
}

function showLocalPathMessageForInput(input) {
  if (input === sheetsUrlInput) {
    if (sheetsHelper) sheetsHelper.textContent = LOCAL_PATH_ERROR_MESSAGE;
    return;
  }
  if (input === apiBaseUrl || input === apiEndpoint) {
    if (apiStatus) apiStatus.textContent = LOCAL_PATH_ERROR_MESSAGE;
    return;
  }
  setStatus(LOCAL_PATH_ERROR_MESSAGE);
}

function validateRemoteInput(value, context = "global") {
  if (!isLikelyLocalWindowsPath(value)) return true;
  if (context === "sheets") {
    if (sheetsHelper) sheetsHelper.textContent = LOCAL_PATH_ERROR_MESSAGE;
  } else if (context === "api") {
    if (apiStatus) apiStatus.textContent = LOCAL_PATH_ERROR_MESSAGE;
  } else {
    setStatus(LOCAL_PATH_ERROR_MESSAGE);
  }
  return false;
}

async function renderLandingDemoPreview() {
  if (!sampleKpiGrid || !sampleTrendChart || !sampleInsightsRoot || !sampleBreakdownChart) return;
  sampleKpiGrid.innerHTML = "<div class=\"helper-text\">Loading demo KPIs...</div>";
  sampleTrendChart.innerHTML = "<div class=\"helper-text\">Loading demo trend...</div>";
  sampleInsightsRoot.innerHTML = "<div class=\"helper-text\">Loading insight...</div>";
  sampleBreakdownChart.innerHTML = "<div class=\"helper-text\">Loading breakdown...</div>";

  try {
    const response = await fetch(DEMO_CSV_URL);
    if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
    const text = await response.text();
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    if (parsed.errors?.length) {
      throw new Error(parsed.errors[0]?.message || "CSV parse error");
    }
    const columns = parsed.meta?.fields || [];
    const rows = (parsed.data || [])
      .filter((row) => row && typeof row === "object")
      .map((row) => {
        const next = {};
        columns.forEach((col) => {
          next[col] = row[col];
        });
        return next;
      });
    if (!rows.length || !columns.length) throw new Error("Demo CSV has no rows.");

    const profiles = profileDataset(rows, columns);
    const { numericCandidates, dimensionCandidates, dateCandidates } = buildSchemaCandidates(columns, profiles);
    const dateDetection = detectDateFieldFromProfiles(rows, profiles);

    landingDemoState.rows = rows;
    landingDemoState.columns = columns;
    landingDemoState.profiles = profiles;
    landingDemoState.numericColumns = (numericCandidates || []).length
      ? numericCandidates
      : columns.filter((col) => profiles[col]?.type === "numeric");
    landingDemoState.categoricalColumns = (dimensionCandidates || []).length
      ? dimensionCandidates
      : columns.filter((col) => profiles[col]?.type === "categorical");
    landingDemoState.dateColumn = dateDetection?.field || dateCandidates?.[0] || null;
    landingDemoState.kpiCards = buildLandingKpiCards();
    landingDemoState.selectedMetric = landingDemoState.kpiCards[0]?.metric || landingDemoState.numericColumns[0] || null;
    landingDemoState.selectedBreakdown = "overall";
    landingDemoState.breakdownMap = {
      plan: chooseLandingDimensionByHint(["plan", "tier", "segment"]),
      channel: chooseLandingDimensionByHint(["channel", "source", "campaign"]),
    };
    landingDemoState.selectedDimension = chooseLandingDimensionByHint(["region", "plan", "channel"])
      || landingDemoState.categoricalColumns[0]
      || null;

    hydrateLandingDemoControls();
    renderLandingKpis();
    renderLandingTrend();
    renderLandingInsightAndBreakdown();
  } catch (error) {
    const message = `Demo preview unavailable: ${error?.message || "unknown error"}`;
    sampleKpiGrid.innerHTML = `<div class="helper-text">${escapeHtml(message)}</div>`;
    sampleTrendChart.innerHTML = `<div class="helper-text">${escapeHtml(message)}</div>`;
    sampleInsightsRoot.innerHTML = `<div class="helper-text">${escapeHtml(message)}</div>`;
    sampleBreakdownChart.innerHTML = `<div class="helper-text">${escapeHtml(message)}</div>`;
  }
}

function hydrateLandingDemoControls() {
  if (sampleTrendMetric) {
    sampleTrendMetric.innerHTML = "";
    const metrics = Array.from(new Set([
      ...landingDemoState.kpiCards.map((card) => card.metric).filter(Boolean),
      ...(landingDemoState.numericColumns || []),
    ]));
    metrics.forEach((metric) => {
      const option = document.createElement("option");
      option.value = metric;
      option.textContent = formatMetricLabel(metric);
      sampleTrendMetric.appendChild(option);
    });
    if (metrics.length) sampleTrendMetric.value = landingDemoState.selectedMetric || metrics[0];
    sampleTrendMetric.onchange = () => {
      landingDemoState.selectedMetric = sampleTrendMetric.value;
      renderLandingTrend();
      renderLandingInsightAndBreakdown();
    };
  }

  if (sampleTrendBreakdown) {
    const planOption = sampleTrendBreakdown.querySelector("option[value='plan']");
    const channelOption = sampleTrendBreakdown.querySelector("option[value='channel']");
    if (planOption) {
      const mappedPlan = landingDemoState.breakdownMap.plan;
      planOption.disabled = !mappedPlan;
      planOption.textContent = mappedPlan ? `By ${formatMetricLabel(mappedPlan)}` : "By plan (not available)";
    }
    if (channelOption) {
      const mappedChannel = landingDemoState.breakdownMap.channel;
      channelOption.disabled = !mappedChannel;
      channelOption.textContent = mappedChannel ? `By ${formatMetricLabel(mappedChannel)}` : "By channel (not available)";
    }
    if (
      sampleTrendBreakdown.value !== "overall"
      && !landingDemoState.breakdownMap[sampleTrendBreakdown.value]
    ) {
      sampleTrendBreakdown.value = "overall";
    }
    landingDemoState.selectedBreakdown = sampleTrendBreakdown.value || "overall";
    sampleTrendBreakdown.onchange = () => {
      landingDemoState.selectedBreakdown = sampleTrendBreakdown.value || "overall";
      renderLandingTrend();
    };
  }

  if (sampleDimensionSelect) {
    sampleDimensionSelect.innerHTML = "";
    const dimensions = (landingDemoState.categoricalColumns || []).slice(0, 6);
    dimensions.forEach((dimension) => {
      const option = document.createElement("option");
      option.value = dimension;
      option.textContent = formatMetricLabel(dimension);
      sampleDimensionSelect.appendChild(option);
    });
    if (dimensions.length) {
      sampleDimensionSelect.value = landingDemoState.selectedDimension || dimensions[0];
      landingDemoState.selectedDimension = sampleDimensionSelect.value;
    }
    sampleDimensionSelect.onchange = () => {
      landingDemoState.selectedDimension = sampleDimensionSelect.value;
      renderLandingInsightAndBreakdown();
    };
  }
}

function chooseLandingDimensionByHint(hints) {
  return (landingDemoState.categoricalColumns || []).find((col) => {
    const lower = String(col || "").toLowerCase();
    return hints.some((hint) => lower.includes(hint));
  }) || null;
}

function buildLandingKpiCards() {
  const targets = [
    { label: "Revenue", hints: ["revenue", "sales", "gmv", "arr", "mrr", "amount"], fallbackKind: "currency" },
    { label: "Retention", hints: ["retention", "returnrate", "rate", "percent", "ratio", "churn"], fallbackKind: "rate" },
    { label: "Active users", hints: ["activeusers", "active_users", "users", "sessions", "visitors", "orders"], fallbackKind: "count" },
    { label: "Pipeline", hints: ["pipeline", "opportunity", "forecast", "backlog", "bookings"], fallbackKind: "currency" },
  ];
  const available = [...(landingDemoState.numericColumns || [])];
  const chosen = [];
  targets.forEach((target) => {
    const metric = pickLandingMetric(target, available, chosen);
    chosen.push(metric);
  });
  return targets.map((target, index) => {
    const metric = chosen[index] || null;
    const metricValues = metric ? landingDemoState.rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null) : [];
    const inferredType = metric ? inferMetricType(metric, metricValues) : { kind: target.fallbackKind };
    const needsLabelSuffix = metric && !target.hints.some((hint) => metric.toLowerCase().includes(hint));
    return {
      label: needsLabelSuffix ? `${target.label} (${formatMetricLabel(metric)})` : target.label,
      metric,
      metricType: inferredType,
    };
  });
}

function pickLandingMetric(target, available, reserved) {
  const candidates = (available || []).filter((metric) => !reserved.includes(metric));
  if (!candidates.length) return null;
  const ranked = candidates
    .map((metric) => {
      const lower = metric.toLowerCase().replace(/[^a-z0-9]/g, "");
      let score = 0;
      target.hints.forEach((hint) => {
        if (lower.includes(hint)) score += hint.length;
      });
      const profile = landingDemoState.profiles?.[metric];
      if (target.fallbackKind === "rate" && profile?.max !== null && profile?.max <= 1.25) score += 3;
      if (target.fallbackKind === "currency" && /revenue|sales|amount|cost|price|pipeline/.test(lower)) score += 2;
      return { metric, score };
    })
    .sort((a, b) => b.score - a.score);
  return ranked[0]?.metric || candidates[0];
}

function getLandingCurrentAndPriorRows(windowDays = 30) {
  const rows = landingDemoState.rows || [];
  const dateColumn = landingDemoState.dateColumn;
  if (!dateColumn) return { currentRows: rows, priorRows: [] };
  const datedRows = rows
    .map((row) => ({ row, date: parseDate(row[dateColumn]) }))
    .filter((entry) => entry.date)
    .sort((a, b) => a.date - b.date);
  if (!datedRows.length) return { currentRows: rows, priorRows: [] };
  const end = datedRows[datedRows.length - 1].date;
  const currentStart = new Date(end);
  currentStart.setUTCDate(currentStart.getUTCDate() - (windowDays - 1));
  const priorEnd = new Date(currentStart);
  priorEnd.setUTCDate(priorEnd.getUTCDate() - 1);
  const priorStart = new Date(priorEnd);
  priorStart.setUTCDate(priorStart.getUTCDate() - (windowDays - 1));
  const currentRows = datedRows
    .filter((entry) => entry.date >= currentStart && entry.date <= end)
    .map((entry) => entry.row);
  const priorRows = datedRows
    .filter((entry) => entry.date >= priorStart && entry.date <= priorEnd)
    .map((entry) => entry.row);
  return { currentRows, priorRows };
}

function aggregateLandingMetric(rows, metric, metricType) {
  if (!metric) return 0;
  const values = (rows || []).map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  if (!values.length) return 0;
  if (metricType?.kind === "rate") {
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }
  return values.reduce((sum, value) => sum + value, 0);
}

function buildLandingKpiSparkline(metric, metricType) {
  const monthly = buildLandingMonthlySeries(metric, null);
  const values = monthly.values || [];
  if (!values.length || values.every((value) => value === 0)) {
    return "<div class=\"helper-text kpi-sparkline-empty\">No time based trend available for this metric.</div>";
  }
  const width = 240;
  const height = 50;
  const padX = 4;
  const padY = 4;
  const minY = Math.min(...values, 0);
  const maxY = Math.max(...values, 0);
  const scaleX = (index) => padX + (index / Math.max(values.length - 1, 1)) * (width - padX * 2);
  const scaleY = (value) => (height - padY) - ((value - minY) / Math.max(maxY - minY, 1)) * (height - padY * 2);
  const path = values.map((value, index) => `${index === 0 ? "M" : "L"}${scaleX(index)},${scaleY(value)}`).join(" ");
  const points = values.map((value, index) => {
    const x = scaleX(index);
    const label = monthly.labels[index] || "";
    return `<circle cx="${x}" cy="${scaleY(value)}" r="4" fill="transparent"><title>${label}: ${formatMetricValue(value, metricType)}</title></circle>`;
  }).join("");
  return `
    <svg class="kpi-sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-label="${escapeHtml(metric)} sparkline">
      <path d="${path}" fill="none" stroke="rgba(96,165,250,0.95)" stroke-width="1.8"></path>
      ${points}
    </svg>
  `;
}

function renderLandingKpis() {
  if (!sampleKpiGrid) return;
  const { currentRows, priorRows } = getLandingCurrentAndPriorRows();
  sampleKpiGrid.innerHTML = "";
  landingDemoState.kpiCards.forEach((cardDef) => {
    const value = aggregateLandingMetric(currentRows, cardDef.metric, cardDef.metricType);
    const priorValue = aggregateLandingMetric(priorRows, cardDef.metric, cardDef.metricType);
    const rawDelta = value - priorValue;
    const direction = rawDelta > 0 ? "up" : rawDelta < 0 ? "down" : "flat";
    const deltaText = cardDef.metricType?.kind === "rate"
      ? `${rawDelta >= 0 ? "+" : ""}${(rawDelta * 100).toFixed(1)} pp`
      : `${rawDelta >= 0 ? "+" : ""}${((priorValue ? rawDelta / Math.abs(priorValue) : 0) * 100).toFixed(1)}%`;
    const sparkline = landingDemoState.dateColumn
      ? buildLandingKpiSparkline(cardDef.metric, cardDef.metricType)
      : "<div class=\"helper-text kpi-sparkline-empty\">No time based trend available for this metric.</div>";
    const card = document.createElement("div");
    card.className = "kpi-pill";
    card.innerHTML = `
      <span>${escapeHtml(cardDef.label)}</span>
      <strong>${formatMetricValue(value, cardDef.metricType)}</strong>
      <em class="kpi-delta ${direction}">${direction === "up" ? "▲" : direction === "down" ? "▼" : "•"} ${deltaText}</em>
      <div class="kpi-mini-trend">${sparkline}</div>
      <div class="kpi-sparkline-caption">Daily trend last 30 days compared to prior 30 days</div>
      <small class="kpi-caption">Last 30 days vs prior 30 days</small>
    `;
    sampleKpiGrid.appendChild(card);
  });
}

function getLandingMonthKey(dateValue) {
  const year = dateValue.getUTCFullYear();
  const month = String(dateValue.getUTCMonth() + 1).padStart(2, "0");
  return `${year}-${month}-01`;
}

function buildLandingRecentMonthKeys(rows, dateColumn, count = 12) {
  const monthKeys = (rows || [])
    .map((row) => parseDate(row[dateColumn]))
    .filter(Boolean)
    .map((dateValue) => getLandingMonthKey(dateValue))
    .sort();
  if (!monthKeys.length) return [];
  const latest = monthKeys[monthKeys.length - 1];
  const latestDate = new Date(`${latest}T00:00:00Z`);
  const keys = [];
  for (let offset = count - 1; offset >= 0; offset -= 1) {
    const monthDate = new Date(Date.UTC(latestDate.getUTCFullYear(), latestDate.getUTCMonth() - offset, 1));
    keys.push(getLandingMonthKey(monthDate));
  }
  return keys;
}

function buildLandingMonthlySeries(metric, breakdownColumn = null) {
  const rows = landingDemoState.rows || [];
  const dateColumn = landingDemoState.dateColumn;
  if (!dateColumn) return { labels: [], groups: [], values: [] };
  const monthKeys = buildLandingRecentMonthKeys(rows, dateColumn, 12);
  if (!monthKeys.length) return { labels: [], groups: [], values: [] };
  const grouped = new Map();
  rows.forEach((row) => {
    const dateValue = parseDate(row[dateColumn]);
    if (!dateValue) return;
    const monthKey = getLandingMonthKey(dateValue);
    if (!monthKeys.includes(monthKey)) return;
    const group = breakdownColumn ? String(row[breakdownColumn] || "Unknown") : "Overall";
    const key = `${group}::${monthKey}`;
    const metricValue = parseNumber(row[metric]);
    if (metricValue === null) return;
    grouped.set(key, (grouped.get(key) || 0) + metricValue);
  });

  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  const groups = breakdownColumn
    ? Array.from(new Set(rows.map((row) => String(row[breakdownColumn] || "Unknown")))).slice(0, 5)
    : ["Overall"];

  const series = groups.map((group) => ({
    group,
    values: monthKeys.map((monthKey) => {
      const total = grouped.get(`${group}::${monthKey}`) || 0;
      if (metricType.kind === "rate") {
        const monthRows = rows.filter((row) => {
          const dateValue = parseDate(row[dateColumn]);
          if (!dateValue || getLandingMonthKey(dateValue) !== monthKey) return false;
          if (!breakdownColumn) return true;
          return String(row[breakdownColumn] || "Unknown") === group;
        });
        const numeric = monthRows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
        return numeric.length ? numeric.reduce((sum, value) => sum + value, 0) / numeric.length : 0;
      }
      return total;
    }),
  }));

  return {
    labels: monthKeys.map((monthKey) => new Date(`${monthKey}T00:00:00Z`).toLocaleDateString("en-US", { month: "short" })),
    rawMonthKeys: monthKeys,
    groups: series,
    values: series[0]?.values || [],
  };
}

function renderLandingTrend() {
  if (!sampleTrendChart) return;
  const metric = landingDemoState.selectedMetric;
  if (!metric) {
    sampleTrendChart.innerHTML = "<div class=\"helper-text\">No numeric metric found in demo CSV.</div>";
    return;
  }
  if (!landingDemoState.dateColumn) {
    sampleTrendChart.innerHTML = `
      <div class="helper-text" style="min-height:220px;display:flex;align-items:center;justify-content:center;">
        No date column found in the demo CSV. Add a date field to render month over month trend.
      </div>
    `;
    return;
  }

  const breakdownField = landingDemoState.selectedBreakdown === "overall"
    ? null
    : landingDemoState.breakdownMap[landingDemoState.selectedBreakdown] || null;
  const series = buildLandingMonthlySeries(metric, breakdownField);
  const allValues = series.groups.flatMap((group) => group.values);
  if (!allValues.length) {
    sampleTrendChart.innerHTML = "<div class=\"helper-text\">No trend data available.</div>";
    return;
  }
  const width = 680;
  const height = 240;
  const left = 30;
  const right = 16;
  const top = 16;
  const bottom = 26;
  const maxY = Math.max(...allValues, 0);
  const minY = Math.min(...allValues, 0);
  const scaleX = (index) => left + (index / Math.max(series.labels.length - 1, 1)) * (width - left - right);
  const scaleY = (value) => (height - bottom) - ((value - minY) / Math.max(maxY - minY, 1)) * (height - top - bottom);
  const colors = ["#60a5fa", "#34d399", "#f59e0b", "#f472b6", "#a78bfa"];
  const grid = [0.25, 0.5, 0.75].map((ratio) => {
    const y = top + ratio * (height - top - bottom);
    return `<line x1="${left}" y1="${y}" x2="${width - right}" y2="${y}" stroke="rgba(255,255,255,0.08)" stroke-width="1"></line>`;
  }).join("");
  const lines = series.groups.map((group, index) => {
    const color = colors[index % colors.length];
    const path = group.values.map((value, pointIndex) => `${pointIndex === 0 ? "M" : "L"}${scaleX(pointIndex)},${scaleY(value)}`).join(" ");
    return `<path d="${path}" fill="none" stroke="${color}" stroke-width="1.8"></path>`;
  }).join("");
  const ticks = series.labels.map((label, index) => `
    <text x="${scaleX(index)}" y="${height - 8}" fill="rgba(255,255,255,0.6)" font-size="10" text-anchor="middle">${label}</text>
  `).join("");
  const legends = series.groups.map((group, index) => `
    <span class="trend-legend-item">
      <span class="trend-legend-swatch" style="background:${colors[index % colors.length]}"></span>
      <span>${escapeHtml(group.group)}</span>
    </span>
  `).join("");
  sampleTrendChart.innerHTML = `
    <div class="trend-legend">${legends}</div>
    <svg class="trend-line" viewBox="0 0 ${width} ${height}" aria-label="Monthly trend chart">
      ${grid}
      ${lines}
      ${ticks}
    </svg>
  `;
}

function computeLandingBreakdown(metric, dimension) {
  if (!metric || !dimension) return [];
  const values = new Map();
  const counts = new Map();
  (landingDemoState.rows || []).forEach((row) => {
    const key = String(row[dimension] || "Unknown");
    const metricValue = parseNumber(row[metric]);
    if (metricValue === null) return;
    values.set(key, (values.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  const metricType = inferMetricType(metric, (landingDemoState.rows || []).map((row) => parseNumber(row[metric])).filter((value) => value !== null));
  return Array.from(values.entries())
    .map(([key, total]) => ({
      key,
      value: metricType.kind === "rate" ? total / Math.max(counts.get(key) || 1, 1) : total,
      count: counts.get(key) || 0,
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);
}

function renderLandingInsightAndBreakdown() {
  const metric = landingDemoState.selectedMetric;
  const dimension = landingDemoState.selectedDimension;
  if (!metric || !dimension) {
    if (sampleInsightsRoot) sampleInsightsRoot.innerHTML = "<div class=\"helper-text\">No categorical dimension found in demo CSV.</div>";
    if (sampleBreakdownChart) sampleBreakdownChart.innerHTML = "<div class=\"helper-text\">No categorical dimension found in demo CSV.</div>";
    return;
  }

  const breakdown = computeLandingBreakdown(metric, dimension);
  if (!breakdown.length) {
    if (sampleInsightsRoot) sampleInsightsRoot.innerHTML = "<div class=\"helper-text\">No breakdown data available.</div>";
    if (sampleBreakdownChart) sampleBreakdownChart.innerHTML = "<div class=\"helper-text\">No breakdown data available.</div>";
    return;
  }

  const metricType = inferMetricType(metric, (landingDemoState.rows || []).map((row) => parseNumber(row[metric])).filter((value) => value !== null));
  const top = breakdown[0];
  const bottom = breakdown[breakdown.length - 1];
  const delta = top.value - bottom.value;
  const deltaLabel = metricType.kind === "rate"
    ? `${(delta * 100).toFixed(1)} pp`
    : formatMetricValue(delta, metricType);
  const confidenceLevel = breakdown.length >= 3 && Math.abs(delta) > 0 ? "high" : "medium";
  const confidenceWhy = confidenceLevel === "high"
    ? "Clear separation between top and bottom segments across enough samples."
    : "Directional difference is visible, but segment count or spread is limited.";
  const driverDimension = (landingDemoState.categoricalColumns || []).find((col) => col !== dimension);
  const driverText = driverDimension
    ? `Driver candidate: ${formatMetricLabel(driverDimension)} appears to concentrate in ${top.key}.`
    : `Driver candidate: Segment mix likely favors ${top.key}.`;
  if (sampleInsightsRoot) {
    sampleInsightsRoot.innerHTML = `
      <div class="insight-card">
        <div class="insight-header">
          <span class="severity ${confidenceLevel} confidence-pill" tabindex="0" aria-label="Confidence details">
            Confidence: ${confidenceLevel === "high" ? "High" : "Medium"}
            <span class="confidence-tooltip" role="tooltip">${escapeHtml(confidenceWhy)}</span>
          </span>
        </div>
        <h4>${escapeHtml(formatMetricLabel(metric))} delta is largest between ${escapeHtml(top.key)} and ${escapeHtml(bottom.key)}</h4>
        <div class="insight-metric">
          <span>Delta</span>
          <strong>${escapeHtml(deltaLabel)}</strong>
        </div>
        <p class="insight-meta">${escapeHtml(driverText)}</p>
        <p class="insight-meta">Action: Replicate the top segment playbook for lower-performing segments and monitor uplift over the next cycle.</p>
        <p class="insight-meta subtle">Compared by ${escapeHtml(formatMetricLabel(dimension))} with n=${top.count}/${bottom.count}.</p>
      </div>
    `;
  }

  if (sampleBreakdownChart) {
    const max = Math.max(...breakdown.map((item) => item.value), 1);
    sampleBreakdownChart.innerHTML = `
      <div class="bar-chart">
        ${breakdown.map((item) => `
          <div class="bar-row">
            <span>${escapeHtml(item.key)}</span>
            <div class="bar"><div class="bar-fill" style="width:${(item.value / max) * 100}%"></div></div>
            <strong>${escapeHtml(formatMetricValue(item.value, metricType))}</strong>
          </div>
        `).join("")}
      </div>
    `;
  }
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
  if (state.uiMode === "ready" && hasLoadedDataset()) {
    console.log("activateSourceTab skipped in ready mode");
    return;
  }
  const sourceType = String(sourceMeta?.sourceType || "").toLowerCase();
  if (sourceMeta?.isSample) {
    switchTab("samples");
    return;
  }
  if (sourceType === "sheet" || sourceType === "sheets") {
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
  if (!validateRemoteInput(url, "sheets")) {
    throw new Error(LOCAL_PATH_ERROR_MESSAGE);
  }
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
  return response.text();
}

async function loadSheetData() {
  const input = sheetsUrlInput?.value || "";
  if (!validateRemoteInput(input, "sheets")) return;
  const range = sheetsRangeInput?.value?.trim() || "A1:Z1000";
  clearMessages();
  const correlationId = createIngestCorrelationId("sheets");
  console.log(`[INGEST ${correlationId}] SHEET selected`, input);
  appStore.setState({ source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "parsing", error: null });
  console.log("SHEET parseStatus -> parsing");
  setUiMode("loading");
  setParseStatus("parsing");
  setStatus("Loading Google Sheet");
  try {
    const spreadsheetId = extractSheetId(input);
    const trimmedInput = input.trim();
    const looksLikeSheetsUrl = /docs\.google\.com\/spreadsheets/i.test(trimmedInput);
    const looksLikeUrl = /^https?:\/\//i.test(trimmedInput);
    if (!trimmedInput || (!spreadsheetId && !looksLikeSheetsUrl && !looksLikeUrl)) {
      const msg = "Paste a valid Google Sheets link or ID to continue.";
      appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: msg });
      setParseStatus("error");
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError(msg);
      return;
    }
    if (googleAccessToken && looksLikeSheetsUrl && !spreadsheetId) {
      const msg = "Invalid Google Sheets link. Make sure it includes the spreadsheet ID.";
      appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: msg });
      setParseStatus("error");
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError(msg);
      return;
    }
    if (googleAccessToken && spreadsheetId) {
      const apiRange = encodeURIComponent(range || "A1:Z1000");
      const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${apiRange}?majorDimension=ROWS`;
      const response = await fetch(apiUrl, {
        headers: { Authorization: `Bearer ${googleAccessToken}` },
      });
      if (response.status === 401 || response.status === 403) {
        updateGoogleStatus(false);
        appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: "Permission denied or token expired. Reconnect Google." });
        console.log("STORE setState parseStatus", appStore.getState().parseStatus);
        showError("Permission denied or token expired. Reconnect Google.");
        return;
      }
      if (!response.ok) {
        const body = await response.text();
        appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: body || "Google Sheets read failed." });
        console.log("STORE setState parseStatus", appStore.getState().parseStatus);
        showError("Google Sheets read failed.", `Details: ${body}`);
        return;
      }
      const payload = await response.json();
      if (!payload?.values || !payload.values.length) {
        appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: "No rows returned from Google Sheets." });
        console.log("STORE setState parseStatus", appStore.getState().parseStatus);
        showError("No rows returned from Google Sheets.");
        return;
      }
      updateGoogleStatus(true);
      console.log(`[INGEST ${correlationId}] SHEET raw rows`, payload.values.length);
      const csvText = valuesToCsv(payload.values || []);
      parseCsvText(csvText, { sourceType: "sheets", name: spreadsheetId || "Google Sheet", correlationId });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      return;
    }

    const normalizedUrl = normalizeSheetsUrl(input);
    if (!normalizedUrl) {
      appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: "Paste a Google Sheets link first." });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError("Paste a Google Sheets link first.");
      return;
    }
    const text = await fetchSheetText(normalizedUrl);
    console.log(`[INGEST ${correlationId}] SHEET csv text length`, text.length);
    parseCsvText(text, { sourceType: "sheets", name: normalizedUrl, correlationId });
    console.log("STORE setState parseStatus", appStore.getState().parseStatus);
  } catch (error) {
    const fallback = buildSheetsFallback(normalizeSheetsUrl(input) || input);
    if (fallback) {
      try {
        const text = await fetchSheetText(fallback);
        console.log(`[INGEST ${correlationId}] SHEET fallback csv text length`, text.length);
        parseCsvText(text, { sourceType: "sheets", name: fallback, correlationId });
        console.log("STORE setState parseStatus", appStore.getState().parseStatus);
        return;
      } catch (fallbackError) {
        console.warn("Fallback CSV fetch failed", fallbackError);
      }
    }
    const detail = error?.message || String(error);
    if (detail.includes("Failed to fetch")) {
      appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: "This sheet is likely private. Publish it or share a public CSV export link." });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError("This sheet is likely private. Publish it or share a public CSV export link.");
    } else {
      appStore.setState({ dataset: null, source: { type: "sheets", name: input || "Google Sheet" }, parseStatus: "error", error: detail });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
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
  const correlationId = createIngestCorrelationId("pdf");
  console.log(`[INGEST ${correlationId}] PDF selected`, file.name, file.size);
  appStore.setState({ source: { type: "pdf", name: file.name }, parseStatus: "parsing", error: null });
  console.log("PDF parseStatus -> parsing");
  setUiMode("loading");
  setParseStatus("parsing");
  renderPdfStatusFromStore();
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
    const totalTablesFound = pageAnalyses.reduce((sum, page) => sum + ((page.tables || []).length), 0);
    console.log("PDF tablesFound", totalTablesFound);

    const selectedTable = selectBestPdfFactTable(pageAnalyses);
    const fallbackRows = pageAnalyses.flatMap((page) => page.rows.map((row) => row.cells || []));
    const finalRows = selectedTable?.rows?.length ? selectedTable.rows : fallbackRows;
    const selectedCols = finalRows?.[0]?.length || 0;
    console.log("PDF selectedTable rows", finalRows.length, "cols", selectedCols);

    if (finalRows.length < 2) {
      const keyValueRows = extractPdfKeyValueRows(pageAnalyses);
      if (keyValueRows.length >= 2) {
        showWarningMessage("Detected key value PDF, showing extracted fields.");
        ingestDataset({
          source: "pdf",
          rows: [["Field", "Value", "Page"], ...keyValueRows.map((row) => [row.Field, row.Value, row.Page])],
          meta: {
            sourceType: "pdf",
            name: file?.name || "PDF upload",
            pdfMode: "text",
            correlationId,
          },
        });
        console.log("STORE setState parseStatus", appStore.getState().parseStatus);
        return;
      }
      const msg = "No structured tables detected in this PDF. Try a report PDF with selectable tables or export as CSV.";
      appStore.setState({
        dataset: null,
        source: { type: "pdf", name: file.name },
        parseStatus: "error",
        error: msg,
      });
      console.log("STORE setState parseStatus", appStore.getState().parseStatus);
      showError(msg);
      if (pdfStatus) pdfStatus.textContent = "PDF ingestion failed";
      return;
    }

    const headerInference = inferPdfHeaders(finalRows);
    let normalizedRows = finalRows;
    if (headerInference.headers.length) {
      normalizedRows = finalRows.slice();
      normalizedRows.splice(headerInference.headerIndex, 1);
      normalizedRows.unshift(headerInference.headers);
    }
    const parsedRowsSample = normalizedRows.slice(0, 3);
    ingestDataset({
      source: "pdf",
      rows: normalizedRows,
      meta: {
        sourceType: "pdf",
        name: file?.name || "PDF upload",
        pdfMode: "table",
        extractedTableKind: selectedTable?.kind || "generic",
        correlationId,
      },
    });
    const headersForDebug = normalizedRows[0] || [];
    const profilesForDebug = profileDataset(normalizedRows.slice(1).map((row) => {
      const obj = {};
      headersForDebug.forEach((header, idx) => {
        obj[header] = row?.[idx] ?? "";
      });
      return obj;
    }), headersForDebug);
    const { numericCandidates, dimensionCandidates } = buildSchemaCandidates(headersForDebug, profilesForDebug);
    console.log({
      detectedHeaders: headersForDebug,
      numericCandidates,
      dimensionCandidates,
      parsedRowsSample,
    });
    console.log("STORE setState parseStatus", appStore.getState().parseStatus);
  } catch (error) {
    const msg = error?.message || "PDF ingestion failed.";
    appStore.setState({
      dataset: null,
      source: { type: "pdf", name: file.name },
      parseStatus: "error",
      error: msg,
    });
    console.log("STORE setState parseStatus", appStore.getState().parseStatus);
    showError("PDF ingestion failed.", `Details: ${msg}`);
    renderPdfStatusFromStore();
  }
}

function extractPdfKeyValueRows(pageAnalyses) {
  const pairs = [];
  const seen = new Set();
  const metricHint = /(total|spend|revenue|sales|roas|ctr|cpc|cpa|clicks|impressions|conversions|pipeline|users|retention|rate)/i;
  const valueHint = /([$€£]?\s*\d[\d,]*(?:\.\d+)?%?)|(\b\d+(?:\.\d+)?\b)/;
  (pageAnalyses || []).forEach((page) => {
    const pageNum = page?.pageNum ?? "";
    (page?.rows || []).forEach((row) => {
      const cells = (row?.cells || []).map((cell) => String(cell || "").trim()).filter(Boolean);
      if (!cells.length) return;
      const joined = cells.join(" ").replace(/\s+/g, " ").trim();
      let field = "";
      let value = "";
      if (cells.length >= 2) {
        field = cells[0];
        value = cells.slice(1).join(" ");
      } else {
        const line = cells[0];
        if (line.includes(":")) {
          const idx = line.indexOf(":");
          field = line.slice(0, idx).trim();
          value = line.slice(idx + 1).trim();
        } else {
          const m = line.match(/^(.+?)\s+([$€£]?\d[\d,]*(?:\.\d+)?%?)$/);
          if (m) {
            field = m[1].trim();
            value = m[2].trim();
          }
        }
      }
      if (!field || !value) return;
      if (!metricHint.test(field) && !metricHint.test(joined)) return;
      if (!valueHint.test(value)) return;
      const key = `${pageNum}|${field}|${value}`;
      if (seen.has(key)) return;
      seen.add(key);
      pairs.push({ Field: field.slice(0, 120), Value: value.slice(0, 120), Page: String(pageNum) });
    });
  });
  return pairs;
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
  if (!validateRemoteInput(base, "api") || !validateRemoteInput(endpoint, "api")) return;
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

function resetStateForNewDataset(options = {}) {
  const preserveStoreState = Boolean(options.preserveStoreState);
  destroyAllCharts();
  state.dataset = null;
  state.filteredRows = [];
  state.schema = { columns: [], profiles: {}, numeric: [], dates: [], categoricals: [] };
  state.numericColumns = [];
  state.debug = { numericCandidates: [], usableNumericMetrics: [], dateCandidates: [] };
  state.selections = { primaryMetric: null, compareMetrics: [] };
  state.selectedMetric = null;
  state.selectedDimension = null;
  state.chartType = "line";
  state.aggregation = "sum";
  state.sort = null;
  state.chartState = null;
  state.industryColumn = null;
  state.chosenXAxisType = "category";
  state.normalizedDataset = null;
  state.dateColumn = null;
  state.dateFieldConfidence = 0;
  state.diagnosticsCorrelation = { xMetric: null, yMetric: null };
  state.previewCompact = true;
  state.uiMode = "empty";
  state.parseStatus = "idle";
  state.error = null;
  if (!preserveStoreState) {
    appStore.setState({
      dataset: null,
      source: null,
      parseStatus: "idle",
      error: null,
      view: "upload",
    });
  }
  updateAnalysisHeaderState(false);
  syncUploadAnalysisState();
  renderDebugUploadStatus();
  if (kpiGrid) kpiGrid.innerHTML = "";
  if (insightsList) insightsList.innerHTML = "";
  if (profileTable) profileTable.innerHTML = "";
  if (table) table.innerHTML = "";
  if (chartsSection) {
    chartsSection.querySelectorAll("canvas").forEach((canvas) => {
      const ctx = canvas.getContext("2d");
      if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
    });
  }
  renderTrendEmptyState("");
  if (dataSummaryPanel) dataSummaryPanel.innerHTML = "";
  if (previewCompactToggle) previewCompactToggle.checked = true;
  if (profileSummaryLine) profileSummaryLine.textContent = "";
  if (diagCorrSummary) diagCorrSummary.textContent = "";
  if (diagCorrTitle) diagCorrTitle.textContent = "Correlation analysis";
  if (diagCorrSubtitle) diagCorrSubtitle.textContent = "r=— • n=—";
  if (diagCorrExplain) diagCorrExplain.textContent = "";
  if (diagCorrTakeaway) diagCorrTakeaway.textContent = "";
  if (diagCorrCaution) diagCorrCaution.textContent = "";
}

function destroyAllCharts() {
  [trendChartInstance, barChartInstance, diagnosticsCorrelationChartInstance, heatmapChartInstance]
    .filter(Boolean)
    .forEach((chart) => chart.destroy());
  trendChartInstance = null;
  barChartInstance = null;
  diagnosticsCorrelationChartInstance = null;
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
  setParseStatus("error", detail || message);
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
  const detection = detectDateFieldFromProfiles(rows, profiles || {});
  return detection.field || null;
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
  const hasDateField = Boolean(inferredDateField);
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
  state.dataset = dataset;
  state.normalizedDataset = dataset;
  state.datasetCapabilities = {
    hasDateField: Boolean(dataset.meta?.hasDateField),
    hasNumericMetrics: Boolean(dataset.meta?.hasNumericMetrics),
    hasCategoricalDimensions: Boolean(dataset.meta?.hasCategoricalDimensions),
  };
  state.mode = mode;
  state.uiMode = dataset?.rows?.length ? "ready" : "empty";
  appStore.setState({
    dataset,
    source: dataset?.meta?.sourceType ? { type: dataset.meta.sourceType, name: dataset.meta?.name || null } : null,
    parseStatus: dataset?.rows?.length ? "ready" : appStore.getState().parseStatus,
    error: null,
    view: dataset?.rows?.length ? "analysis" : "upload",
  });
  if (dataset?.rows?.length) setParseStatus("ready");
  dashboard?.setAttribute("data-dashboard-mode", mode || "user");
  dashboard?.setAttribute("data-source-type", dataset.meta?.sourceType || "csv");
  if (dataset.meta && dataset.meta.hasDateField === false) {
    state.dateColumn = null;
  } else if (dataset.meta?.dateField) {
    state.dateColumn = dataset.meta.dateField;
  }
  state.chartType = state.dateColumn ? "line" : "bar";
  syncChartStateDirect({
    timeField: state.dateColumn,
    chartType: state.chartType,
    selectedMetric: state.selections.primaryMetric,
    selectedDimension: state.selectedDimension,
    topN: state.topN,
  });
  ensureUnifiedUserDashboardLayout();
  ensureDashboardActionsToolbar();
  dashboard?.classList.add("dashboard-unified-active");
  updateAnalysisHeaderState(Boolean(dataset?.rows?.length));
  syncUploadAnalysisState();
  applyFiltersAndRender();
}

function ingestDataset({ source, rows, columns, meta = {} }) {
  const sourceType = String(source || meta?.sourceType || "csv").toLowerCase();
  const sourceName = meta?.name || null;
  if (meta?.correlationId) {
    console.log(`[INGEST ${meta.correlationId}] ingestDataset`, sourceType, sourceName || "");
  }
  appStore.setState({
    dataset: null,
    source: { type: sourceType, name: sourceName },
    parseStatus: "parsing",
    error: null,
    view: "upload",
  });
  setUiMode("loading");
  setParseStatus("parsing");
  if (Array.isArray(rows) && rows.length && Array.isArray(rows[0])) {
    ingestRows(rows, { ...meta, sourceType });
    return;
  }
  if (Array.isArray(rows) && rows.length && typeof rows[0] === "object" && !Array.isArray(rows[0])) {
    const inferredColumns = Array.isArray(columns) && columns.length
      ? columns
      : Array.from(rows.reduce((set, row) => {
        Object.keys(row || {}).forEach((key) => set.add(key));
        return set;
      }, new Set()));
    const rawRows = [inferredColumns, ...rows.map((row) => inferredColumns.map((col) => row?.[col] ?? ""))];
    ingestRows(rawRows, { ...meta, sourceType });
    return;
  }
  appStore.setState({
    dataset: null,
    source: { type: sourceType, name: sourceName },
    parseStatus: "error",
    error: "No rows found to ingest.",
    view: "upload",
  });
  showError("No rows found to ingest.");
}

function ingestRows(rawRows, sourceMeta = {}) {
  try {
  setStatus("Profiling columns");
  setParseStatus("parsing");
  if (sourceMeta?.correlationId) {
    console.log(`[INGEST ${sourceMeta.correlationId}] ingestRows start`, rawRows?.length || 0);
  }
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

  resetStateForNewDataset({ preserveStoreState: true });
  appStore.setState({
    dataset: null,
    source: {
      type: sourceMeta?.sourceType || "csv",
      name: sourceMeta?.name || null,
    },
    parseStatus: "parsing",
    error: null,
    view: "upload",
  });
  setUiMode("loading");
  state.rawRows = rows;
  state.schema.columns = cleanedHeaders;
  state.schema.profiles = profileDataset(rows, cleanedHeaders);
  state.inferredDomain = inferDomain(state.schema.profiles);
  state.domain = state.domainAuto ? state.inferredDomain : state.domain;
  const { numericCandidates, dimensionCandidates, dateCandidates } = buildSchemaCandidates(state.schema.columns, state.schema.profiles);
  const usableNumericMetrics = numericCandidates.filter((col) => {
    let countParsed = 0;
    for (const row of state.rawRows) {
      if (parseNumber(row[col]) !== null) {
        countParsed += 1;
        if (countParsed >= 1) return true;
      }
    }
    return false;
  }).filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], rows.length));
  const categoricalCandidates = dimensionCandidates.filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], rows.length));
  state.schema.numeric = Array.isArray(usableNumericMetrics) ? usableNumericMetrics : [];
  state.numericColumns = state.schema.numeric;
  state.schema.dates = Array.isArray(dateCandidates) ? dateCandidates : [];
  state.schema.categoricals = Array.isArray(categoricalCandidates) ? categoricalCandidates : [];
  state.debug.numericCandidates = numericCandidates;
  state.debug.usableNumericMetrics = state.schema.numeric;
  state.debug.dateCandidates = state.schema.dates;
  state.industryColumn = findIndustryColumn(state.schema.columns);
  const dateDetection = detectDateFieldFromProfiles(rows, state.schema.profiles);
  state.dateColumn = dateDetection.field;
  state.dateFieldConfidence = dateDetection.confidence;
  syncChartStateDirect({ timeField: state.dateColumn });

  ensureMetrics(state.schema, state.rawRows);

  const recommendedMetrics = chooseKpiMetrics(state.schema.profiles, state.schema.numeric);
  state.selections.primaryMetric = recommendedMetrics[0] || state.schema.numeric[0] || null;
  state.selectedMetric = state.selections.primaryMetric;
  state.selectedDimension = state.selectedDimension
    || dimensionCandidates[0]
    || chooseBestDimension(state.schema.profiles, state.schema.categoricals)
    || state.schema.dates[0]
    || null;
  state.chartType = state.dateColumn ? "line" : "bar";
  syncChartStateDirect({
    selectedMetric: state.selections.primaryMetric,
    selectedDimension: state.selectedDimension,
    chartType: state.chartType,
    topN: state.topN,
    aggregation: state.aggregation,
  });
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
  console.log("NORMALIZE result", state.normalizedDataset?.rows?.length || 0, state.normalizedDataset?.columns?.length || 0);

  initControls(recommendedMetrics);
  setStatus("Rendering dashboard");
  renderDashboardRenderer({
    dataset: state.normalizedDataset,
    mode: state.mode,
  });

  const sourceBadge = state.normalizedDataset?.meta?.sourceType ? ` · ${String(state.normalizedDataset.meta.sourceType).toUpperCase()}` : "";
  state.dataset = state.normalizedDataset;
  console.log("STATE dataset set", state.dataset?.rows?.length || 0);
  appStore.setState({
    dataset: state.dataset,
    source: (state.dataset?.meta?.sourceType || sourceMeta?.sourceType)
      ? {
          type: state.dataset?.meta?.sourceType || sourceMeta?.sourceType,
          name: state.dataset?.meta?.name || sourceMeta?.name || null,
        }
      : null,
    parseStatus: "ready",
    error: null,
    view: "analysis",
  });
  datasetSummary.textContent = `${rows.length} rows · ${state.schema.columns.length} columns${sourceBadge}`;
  activateSourceTab(sourceMeta);
  onDatasetReady(sourceMeta);
  if (statusSection) {
    statusSection.classList.add("hidden");
    statusSection.textContent = "";
  }
  setUiMode(state.dataset?.rows?.length ? "ready" : "empty");
  console.log("dashboard rendered");
  } catch (error) {
    console.error("INGEST rows failed", error);
    appStore.setState({
      dataset: null,
      source: {
        type: sourceMeta?.sourceType || appStore.getState().source?.type || "csv",
        name: sourceMeta?.name || appStore.getState().source?.name || null,
      },
      parseStatus: "error",
      error: error?.message || String(error),
      view: "upload",
    });
    showError("Data ingestion failed while building the dashboard.", `Details: ${error?.message || error}`);
  }
}

function createIngestCorrelationId(sourceType = "data") {
  const stamp = Date.now();
  const rand = Math.random().toString(36).slice(2, 8);
  return `${sourceType}-${stamp}-${rand}`;
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

function getSemanticRole(columnName) {
  const key = String(columnName || "").toLowerCase();
  if (key.includes("cpm")) return "metric_cpm";
  if (key.includes("ctr")) return "metric_ctr";
  if (key.includes("revenue")) return "metric_revenue";
  if (key.includes("date")) return "dimension_date";
  if (key.includes("region")) return "dimension_region";
  return null;
}

function buildSchemaCandidates(columns, profiles) {
  const numericCandidates = [];
  const dimensionCandidates = [];
  const dateCandidates = [];

  (columns || []).forEach((col) => {
    const profile = profiles?.[col];
    const numericRate = profile?.numericRate ?? 0;
    const uniqueCount = profile?.uniqueCount ?? 0;
    const role = getSemanticRole(col);

    if (profile?.type === "date") {
      dateCandidates.push(col);
    }
    if (numericRate >= 0.6 && profile?.type === "numeric") {
      numericCandidates.push(col);
    }
    if (role && role.startsWith("metric_") && numericRate >= 0.6 && profile?.type === "numeric") {
      if (!numericCandidates.includes(col)) numericCandidates.push(col);
    }
    if ((numericRate < 0.2 && uniqueCount > 1) || (role && role.startsWith("dimension_"))) {
      dimensionCandidates.push(col);
    }
  });

  return { numericCandidates, dimensionCandidates, dateCandidates };
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

  domainSelect.value = state.domainAuto ? "auto" : state.domain;
}

function applyFiltersAndRender() {
  clearMessages();
  const scopedRows = Array.isArray(state.rawRows) ? [...state.rawRows] : [];
  state.filteredRows = scopedRows;
  if (!Array.isArray(scopedRows) || scopedRows.length === 0) {
    renderDataSummary([]);
    showError("No data available.");
    renderTrendEmptyState("No data available.");
    return;
  }

  state.schema.profiles = profileDataset(scopedRows, state.schema.columns);
  const { numericCandidates, dimensionCandidates, dateCandidates } = buildSchemaCandidates(state.schema.columns, state.schema.profiles);
  const usableNumeric = (numericCandidates || []).filter((col) => {
    let countParsed = 0;
    for (const row of scopedRows) {
      if (parseNumber(row[col]) !== null) {
        countParsed += 1;
        if (countParsed >= 1) return true;
      }
    }
    return false;
  }).filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], scopedRows.length));
  const categoricalCandidates = dimensionCandidates.filter((col) => !isIdLikeColumn(col, state.schema.profiles[col], scopedRows.length));
  state.schema.numeric = Array.isArray(usableNumeric) ? usableNumeric : [];
  state.numericColumns = state.schema.numeric;
  state.schema.dates = Array.isArray(dateCandidates) ? dateCandidates : [];
  state.schema.categoricals = Array.isArray(categoricalCandidates) ? categoricalCandidates : [];
  state.debug.numericCandidates = numericCandidates;
  state.debug.usableNumericMetrics = state.schema.numeric;
  state.debug.dateCandidates = state.schema.dates;
  const dateDetection = detectDateFieldFromProfiles(scopedRows, state.schema.profiles);
  state.dateColumn = dateDetection.field;
  state.dateFieldConfidence = dateDetection.confidence;
  syncChartStateDirect({ timeField: state.dateColumn });
  ensureMetrics(state.schema, scopedRows);
  if (!state.selections.primaryMetric || !state.schema.numeric.includes(state.selections.primaryMetric)) {
    state.selections.primaryMetric = state.schema.numeric[0] || null;
  }
  state.selectedMetric = state.selections.primaryMetric;
  state.selections.compareMetrics = (state.selections.compareMetrics || []).filter((m) => state.schema.numeric.includes(m));
  if (!state.selectedDimension) {
    state.selectedDimension = dimensionCandidates[0] || chooseBestDimension(state.schema.profiles, state.schema.categoricals) || state.schema.dates[0] || null;
  }
  syncChartStateDirect({
    selectedMetric: state.selections.primaryMetric,
    selectedDimension: state.selectedDimension,
    topN: state.topN,
  });

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

  if (datasetInlineNotice) {
    if (!state.schema.dates.length) {
      datasetInlineNotice.textContent = "No date field detected. Showing distribution view.";
      datasetInlineNotice.classList.remove("hidden");
    } else {
      datasetInlineNotice.classList.add("hidden");
      datasetInlineNotice.textContent = "";
    }
  }

  if (!runStep("buildKPIs", () => renderKPIs(scopedRows))) return;
  if (!runStep("buildCharts", () => renderCharts(scopedRows))) return;
  if (!runStep("buildDiagnosticsCorrelation", () => renderDiagnosticsCorrelation(scopedRows))) return;
  if (!runStep("buildTable", () => renderTable(scopedRows, state.schema.columns))) return;
  if (!runStep("buildInsights", () => renderInsights(scopedRows))) return;
  runStep("buildProfile", () => renderProfileTable(state.schema.profiles));
  runStep("buildQuality", () => renderQualityBadge(scopedRows));
  runStep("buildWarnings", () => showWarnings(collectWarnings()));
  runStep("buildSummary", () => renderDataSummary(scopedRows));
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

function detectDateFieldFromProfiles(rows, profiles) {
  if (detectDateFieldModule?.detectDateField) {
    return detectDateFieldModule.detectDateField({
      columns: state.schema.columns,
      profiles,
      rows,
    });
  }
  return { field: chooseBestDateColumn(profiles), confidence: 0.4 };
}

function getAnalysisRows() {
  return Array.isArray(state.rawRows) ? state.rawRows : [];
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
      : `<div class="kpi-sparkline-caption">Snapshot metric only.</div>`;

    const card = document.createElement("div");
    card.className = "kpi-card user-kpi-card";
    card.innerHTML = `
      <h4>${metricLabel}</h4>
      <div class="kpi-value">${formatMetricValue(primary, metricType)}</div>
      ${deltaBadge}
      ${sparkline}
      <div class="kpi-caption">${supportsTrend ? "Last 30 days vs prior 30 days" : "Snapshot metric only."}</div>
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

function renderTrendEmptyState(message) {
  const trendCanvas = document.getElementById("trendChart");
  if (!trendCanvas || !trendEmptyState) return;
  const text = String(message || "").trim();
  trendEmptyState.classList.remove("trend-rich");
  if (!text) {
    trendEmptyState.classList.add("hidden");
    trendEmptyState.innerHTML = "";
    trendCanvas.classList.remove("hidden");
    return;
  }
  trendEmptyState.textContent = text;
  trendEmptyState.classList.remove("hidden");
  trendCanvas.classList.add("hidden");
}

function renderTrendFallbackContent(html) {
  const trendCanvas = document.getElementById("trendChart");
  if (!trendCanvas || !trendEmptyState) return;
  trendEmptyState.innerHTML = html || "";
  trendEmptyState.classList.add("trend-rich");
  trendEmptyState.classList.remove("hidden");
  trendCanvas.classList.add("hidden");
}

function scheduleTrendChartResize() {
  if (!trendChartInstance) return;
  requestAnimationFrame(() => {
    if (!trendChartInstance) return;
    trendChartInstance.resize();
    trendChartInstance.update("none");
  });
}

function renderDataSummary(rows) {
  if (!dataSummaryPanel) return;
  const rowCount = rows?.length || 0;
  const metricCount = state.schema.numeric?.length || 0;
  const dimensionCount = state.schema.categoricals?.length || 0;
  const dateProfile = state.dateColumn ? state.schema.profiles?.[state.dateColumn] : null;
  const dateRangeText = dateProfile?.dateMin && dateProfile?.dateMax
    ? `${formatDateInput(dateProfile.dateMin)} → ${formatDateInput(dateProfile.dateMax)}`
    : "Unavailable";
  dataSummaryPanel.innerHTML = `
    <div class="summary-item"><strong>Rows</strong><span>${numberFormatter.format(rowCount)}</span></div>
    <div class="summary-item"><strong>Date range</strong><span>${dateRangeText}</span></div>
    <div class="summary-item"><strong>Dimensions</strong><span>${dimensionCount}</span></div>
    <div class="summary-item"><strong>Metrics</strong><span>${metricCount}</span></div>
  `;
}

function detectDateField(rows) {
  if (state.dateColumn) return state.dateColumn;
  const detection = detectDateFieldFromProfiles(rows || [], state.schema.profiles || {});
  return detection?.field || null;
}

function getTimeBuckets(rows, dateField) {
  if (!dateField) return { bucket: "day", keys: [] };
  const bucket = rows.length > 200 ? "week" : "day";
  const keys = new Set();
  (rows || []).forEach((row) => {
    const dateValue = parseDate(row?.[dateField]);
    if (!dateValue) return;
    const key = bucket === "week" ? getWeekStart(dateValue) : formatDateInput(dateValue);
    keys.add(key);
  });
  return { bucket, keys: Array.from(keys).sort() };
}

function pickHeatmapDims(schema) {
  const rowPriority = ["region", "state", "country", "campaign", "line item"];
  const colPriority = ["device", "supply type", "channel", "creative size"];
  const categorical = (schema?.categoricals || []).filter((col) => {
    const unique = schema?.profiles?.[col]?.uniqueCount || 0;
    return unique >= 2 && unique <= 60;
  });
  if (categorical.length < 2) return null;
  const findByPriority = (priority, excluded = null) => {
    for (const hint of priority) {
      const found = categorical.find((col) => {
        if (excluded && col === excluded) return false;
        const normalized = String(col || "").toLowerCase().replace(/[_-]/g, " ");
        return normalized.includes(hint);
      });
      if (found) return found;
    }
    return categorical.find((col) => !excluded || col !== excluded) || null;
  };
  const rowDim = findByPriority(rowPriority);
  const colDim = findByPriority(colPriority, rowDim);
  if (!rowDim || !colDim || rowDim === colDim) return null;
  return { rowDim, colDim };
}

function buildCappedMatrix(rows, rowDim, colDim, metric, maxRows = MAX_HEATMAP_ROWS, maxCols = MAX_HEATMAP_COLS) {
  const metricValues = (rows || []).map((row) => parseNumber(row?.[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  const denominator = metricType.kind === "rate" ? findRateDenominator(metric, state.schema.numeric || []) : null;
  const cellMap = new Map();
  const rowTotals = new Map();
  const colTotals = new Map();
  const rawRecords = [];

  const getCell = (r, c) => cellMap.get(`${r}::${c}`) || {
    row: r, col: c, sum: 0, count: 0, weightedSum: 0, weightSum: 0, contribution: 0,
  };
  const setCell = (cell) => cellMap.set(`${cell.row}::${cell.col}`, cell);

  (rows || []).forEach((row) => {
    const rowKey = String(row?.[rowDim] ?? "Unknown").trim() || "Unknown";
    const colKey = String(row?.[colDim] ?? "Unknown").trim() || "Unknown";
    const metricValue = parseNumber(row?.[metric]);
    if (metricValue === null) return;
    const weight = denominator ? Math.max(0, parseNumber(row?.[denominator]) || 0) : 1;
    const contribution = metricType.kind === "rate" ? metricValue * weight : metricValue;
    rawRecords.push({ rowKey, colKey, metricValue, weight });
    rowTotals.set(rowKey, (rowTotals.get(rowKey) || 0) + contribution);
    colTotals.set(colKey, (colTotals.get(colKey) || 0) + contribution);
    const cell = getCell(rowKey, colKey);
    cell.sum += metricValue;
    cell.count += 1;
    cell.weightedSum += metricValue * weight;
    cell.weightSum += weight;
    cell.contribution += contribution;
    setCell(cell);
  });
  if (!rawRecords.length) return null;

  const sortedRows = Array.from(rowTotals.entries()).sort((a, b) => b[1] - a[1]).map(([key]) => key);
  const sortedCols = Array.from(colTotals.entries()).sort((a, b) => b[1] - a[1]).map(([key]) => key);
  const topRows = sortedRows.slice(0, Math.max(1, maxRows - 1));
  const topCols = sortedCols.slice(0, Math.max(1, maxCols - 1));
  const rowSet = new Set(topRows);
  const colSet = new Set(topCols);

  const cappedRows = [...topRows];
  const cappedCols = [...topCols];
  if (sortedRows.length > topRows.length) cappedRows.push("Other");
  if (sortedCols.length > topCols.length) cappedCols.push("Other");

  const cappedMap = new Map();
  const getCapped = (r, c) => cappedMap.get(`${r}::${c}`) || {
    row: r, col: c, sum: 0, count: 0, weightedSum: 0, weightSum: 0, contribution: 0,
  };
  const setCapped = (cell) => cappedMap.set(`${cell.row}::${cell.col}`, cell);
  rawRecords.forEach((record) => {
    const mappedRow = rowSet.has(record.rowKey) ? record.rowKey : "Other";
    const mappedCol = colSet.has(record.colKey) ? record.colKey : "Other";
    const cell = getCapped(mappedRow, mappedCol);
    const contribution = metricType.kind === "rate" ? record.metricValue * record.weight : record.metricValue;
    cell.sum += record.metricValue;
    cell.count += 1;
    cell.weightedSum += record.metricValue * record.weight;
    cell.weightSum += record.weight;
    cell.contribution += contribution;
    setCapped(cell);
  });

  const materializeCells = (rowsList, colsList, map) => {
    let totalContribution = 0;
    const cells = rowsList.flatMap((rowKey) => colsList.map((colKey) => {
      const raw = map.get(`${rowKey}::${colKey}`) || { sum: 0, count: 0, weightedSum: 0, weightSum: 0, contribution: 0 };
      const value = metricType.kind === "rate"
        ? (raw.weightSum > 0 ? raw.weightedSum / raw.weightSum : (raw.count ? raw.sum / raw.count : 0))
        : raw.sum;
      totalContribution += raw.contribution;
      return { row: rowKey, col: colKey, value, contribution: raw.contribution };
    }));
    return { cells, totalContribution: totalContribution || 1 };
  };

  const capped = materializeCells(cappedRows, cappedCols, cappedMap);
  const full = materializeCells(sortedRows, sortedCols, cellMap);
  return {
    rowDim,
    colDim,
    metric,
    metricType,
    cappedRows,
    cappedCols,
    cappedCells: capped.cells,
    cappedTotalContribution: capped.totalContribution,
    fullRows: sortedRows,
    fullCols: sortedCols,
    fullCells: full.cells,
    fullTotalContribution: full.totalContribution,
  };
}

function formatHeatmapCellValue(value, metricType) {
  if (metricType?.kind === "rate") return `${(Number(value || 0) * 100).toFixed(1)}%`;
  if (metricType?.kind === "currency") return compactCurrencyFormatter.format(Number(value || 0));
  return compactNumberFormatter.format(Number(value || 0));
}

function openTrendHeatmapDrilldown(matrix) {
  if (!matrix || !evidenceDrawer || !evidenceContent) return;
  const rows = matrix.fullRows || [];
  const cols = matrix.fullCols || [];
  const lookup = new Map((matrix.fullCells || []).map((cell) => [`${cell.row}::${cell.col}`, cell]));
  const header = cols.map((col) => `<th title="${escapeHtml(col)}">${escapeHtml(truncateLabel(col, 22))}</th>`).join("");
  const body = rows.map((row) => {
    const cells = cols.map((col) => {
      const cell = lookup.get(`${row}::${col}`) || { value: 0, contribution: 0 };
      const share = (cell.contribution / Math.max(matrix.fullTotalContribution || 1, 1)) * 100;
      const tooltip = `${row} \u00d7 ${col} | ${formatMetricValue(cell.value, matrix.metricType)} | Share ${share.toFixed(1)}%`;
      return `<td title="${escapeHtml(tooltip)}">${escapeHtml(formatHeatmapCellValue(cell.value, matrix.metricType))}</td>`;
    }).join("");
    return `<tr><th title="${escapeHtml(row)}">${escapeHtml(truncateLabel(row, 26))}</th>${cells}</tr>`;
  }).join("");
  evidenceDrawer.classList.remove("hidden");
  evidenceDrawer.setAttribute("aria-hidden", "false");
  if (evidenceScrim) {
    evidenceScrim.classList.remove("hidden");
    evidenceScrim.setAttribute("aria-hidden", "false");
  }
  evidenceContent.innerHTML = `
    <div class="evidence-section">
      <strong>Full breakdown (${escapeHtml(matrix.rowDim)} \u00d7 ${escapeHtml(matrix.colDim)})</strong>
      <div class="helper-text">Scrollable full matrix drilldown. Export uses the capped matrix only.</div>
    </div>
    <div class="evidence-table-wrap">
      <table class="evidence-table">
        <thead><tr><th>${escapeHtml(matrix.rowDim)}</th>${header}</tr></thead>
        <tbody>${body || `<tr><td colspan="${cols.length + 1}">No rows available.</td></tr>`}</tbody>
      </table>
    </div>
  `;
}

function renderHeatmap(matrix, container) {
  if (!matrix || !container) return;
  const cells = matrix.cappedCells || [];
  const maxValue = Math.max(...cells.map((cell) => cell.value), 0);
  const minValue = Math.min(...cells.map((cell) => cell.value), 0);
  const totalGridRows = Math.max(1, (matrix.cappedRows || []).length + 1);
  const computedCellHeight = Math.max(20, Math.min(34, Math.floor(232 / totalGridRows)));
  const scaleColor = (value) => {
    const t = (value - minValue) / Math.max(maxValue - minValue, 1);
    const alpha = 0.22 + (t * 0.72);
    return `rgba(139, 92, 246, ${alpha.toFixed(3)})`;
  };
  const colHeaders = (matrix.cappedCols || [])
    .map((col) => `<div class="heatmapCell heatmapHead" title="${escapeHtml(col)}">${escapeHtml(truncateLabel(col, 14))}</div>`)
    .join("");
  const rowsMarkup = (matrix.cappedRows || []).map((row) => {
    const rowHeader = `<div class="heatmapCell heatmapHead" title="${escapeHtml(row)}">${escapeHtml(truncateLabel(row, 16))}</div>`;
    const rowCells = (matrix.cappedCols || []).map((col) => {
      const cell = cells.find((entry) => entry.row === row && entry.col === col) || { value: 0, contribution: 0 };
      const share = (cell.contribution / Math.max(matrix.cappedTotalContribution || 1, 1)) * 100;
      const tooltip = `${row} \u00d7 ${col}\n${formatMetricValue(cell.value, matrix.metricType)}\nShare: ${share.toFixed(1)}%`;
      return `<div class="heatmapCell" title="${escapeHtml(tooltip)}" style="background:${scaleColor(cell.value)}">${escapeHtml(formatHeatmapCellValue(cell.value, matrix.metricType))}</div>`;
    }).join("");
    return `${rowHeader}${rowCells}`;
  }).join("");

  renderTrendFallbackContent(`
    <div class="hmHeader">
      <div class="heatmapCaption helper-text">Showing Top ${MAX_HEATMAP_ROWS} rows \u00d7 Top ${MAX_HEATMAP_COLS} columns. Remaining grouped as Other.</div>
    </div>
    <div class="hmLegend heatmapLegend"><span>Low</span><div class="heatmapLegendBar"></div><span>High</span></div>
    <div class="hmGridWrap">
      <div class="heatmapGrid" style="--hm-cell-h:${computedCellHeight}px;grid-template-columns:minmax(70px,1.05fr) repeat(${Math.max(1, (matrix.cappedCols || []).length)}, minmax(0,1fr));">
        <div class="heatmapCell heatmapHead"></div>
        ${colHeaders}
        ${rowsMarkup}
      </div>
    </div>
    <div class="hmFooter">
      <button type="button" id="trendHeatmapDrilldown" class="ghost trend-drilldown">View full breakdown</button>
    </div>
  `);
  const drilldownBtn = document.getElementById("trendHeatmapDrilldown");
  if (drilldownBtn) {
    drilldownBtn.addEventListener("click", () => openTrendHeatmapDrilldown(matrix));
  }
}

function renderRankedBars(rows, dim, metric, container) {
  const metricValues = (rows || []).map((row) => parseNumber(row?.[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  if (!dim) {
    renderTrendEmptyState("No data available.");
    return;
  }
  const ranked = buildTopCategoriesWithOther(rows, dim, metric, metricType, 8);
  const hasValues = (ranked.values || []).some((value) => Number(value) > 0);
  if (!hasValues || !(ranked.labels || []).length) {
    renderTrendEmptyState("No measurable activity for selected dimension.");
    return;
  }
  renderTrendFallbackContent(`
    <div class="heatmapCaption helper-text">Time trend unavailable. Showing ranked bars by ${escapeHtml(dim)}.</div>
    <div class="bar-chart">
      ${ranked.labels.map((label, idx) => {
        const max = Math.max(...ranked.values, 1);
        const width = ((ranked.values[idx] || 0) / max) * 100;
        return `
          <div class="bar-row" title="${escapeHtml(label)}">
            <span>${escapeHtml(truncateLabel(label, 20))}</span>
            <div class="bar"><div class="bar-fill" style="width:${width}%"></div></div>
            <strong>${escapeHtml(formatMetricValue(ranked.values[idx], metricType))}</strong>
          </div>
        `;
      }).join("")}
    </div>
  `);
}

function renderTrendModule(data, selectedMetric) {
  const rows = Array.isArray(data) ? data : [];
  const metric = selectedMetric;
  const metricValues = rows.map((row) => parseNumber(row?.[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  if (!rows.length) {
    trendTitle.textContent = "Trend";
    trendSubtitle.textContent = "No data available";
    renderTrendEmptyState("No data available.");
    return;
  }

  const dateField = detectDateField(rows);
  const timeInfo = dateField ? getTimeBuckets(rows, dateField) : { bucket: "day", keys: [] };
  if (dateField && timeInfo.keys.length >= 2) {
    state.chosenXAxisType = "date";
    const series = aggregateByDate(rows, dateField, metric, timeInfo.bucket, null, metricType);
    if ((series?.labels || []).length >= 2) {
      renderTrendEmptyState("");
      const desiredChartType = state.chartType || "line";
      trendTitle.textContent = "Trend";
      const periodLabel = series.labels.length < 3 ? `${series.labels.length} periods (limited)` : `${series.labels.length} periods`;
      trendSubtitle.textContent = `${formatMetricLabel(metric)} over time · ${periodLabel}`;
      trendChartInstance = desiredChartType === "bar"
        ? createBarChart("trendChart", series.labels, series.values, metricType, series.counts)
        : createLineChart("trendChart", series.labels, series.values, metricType, series.counts);
      scheduleTrendChartResize();
      return;
    }
  }

  const heatmapDims = pickHeatmapDims(state.schema);
  if (heatmapDims) {
    const matrix = buildCappedMatrix(rows, heatmapDims.rowDim, heatmapDims.colDim, metric, MAX_HEATMAP_ROWS, MAX_HEATMAP_COLS);
    if (matrix && matrix.cappedCells.length) {
      trendTitle.textContent = "Breakdown";
      trendSubtitle.textContent = `${formatMetricLabel(heatmapDims.rowDim)} \u00d7 ${formatMetricLabel(heatmapDims.colDim)} heatmap for ${formatMetricLabel(metric)}`;
      renderHeatmap(matrix, trendEmptyState);
      return;
    }
  }

  trendTitle.textContent = "Trend";
  trendSubtitle.textContent = "Ranked overview";
  const fallbackDim = chooseBestDimension(state.schema.profiles, state.schema.categoricals || [])
    || chooseCategoryColumn(rows, state.schema.profiles);
  renderRankedBars(rows, fallbackDim, metric, trendEmptyState);
}

function renderCharts(rows) {
  if (trendChartInstance) trendChartInstance.destroy();
  if (barChartInstance) barChartInstance.destroy();

  const metric = state.selections.primaryMetric || chooseTopNumericByVariance(state.schema.profiles, state.schema.numeric || []);
  state.selections.primaryMetric = metric;
  state.selectedMetric = metric;
  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);

  renderTrendModule(rows, metric);

  let dimension = state.selectedDimension || chooseBestDimension(state.schema.profiles, state.schema.categoricals || []) || state.dateColumn;
  if (dimension) {
    state.selectedDimension = dimension;
    if (dimensionSelect.value !== dimension) dimensionSelect.value = dimension;
    const breakdown = buildTopCategoriesWithOther(rows, dimension, metric, metricType, 8);
    barTitle.textContent = "Performance details";
      barSubtitle.textContent = `Dimension: ${dimension}`;
      const hasMeasuredValues = (breakdown.values || []).some((value) => Number(value) > 0);
      if (!breakdown.labels.length || !hasMeasuredValues) {
        setChartCardEmptyState("barChart", "No measurable activity for selected dimension.");
      } else {
        setChartCardEmptyState("barChart", "");
        barChartInstance = createHorizontalBarChart("barChart", breakdown.labels, breakdown.values, metricType, breakdown.counts);
      }
  } else {
    barTitle.textContent = "Performance details";
    barSubtitle.textContent = "Dimension unavailable";
    setChartCardEmptyState("barChart", "Dimension unavailable in dataset.");
    state.selectedDimension = null;
    if (dimensionSelect) dimensionSelect.value = "";
  }
}

function renderProfileTable(profile) {
  if (profileSummaryLine) {
    const rowsCount = numberFormatter.format((state.filteredRows || []).length);
    const colsCount = numberFormatter.format((state.schema.columns || []).length);
    const timeField = state.dateColumn ? formatMetricLabel(state.dateColumn) : "none";
    profileSummaryLine.textContent = `Rows: ${rowsCount} • Columns: ${colsCount} • Detected time field: ${timeField}`;
  }
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

function getCompactPreviewColumns(columns) {
  const available = Array.isArray(columns) ? [...columns] : [];
  const chosen = [];
  const pushUnique = (col) => {
    if (!col || !available.includes(col) || chosen.includes(col)) return;
    chosen.push(col);
  };
  pushUnique(state.dateColumn);
  pushUnique(state.selectedDimension);
  pushUnique(state.selections.primaryMetric);
  (state.schema.numeric || []).slice(0, 4).forEach((col) => pushUnique(col));
  (state.schema.categoricals || []).slice(0, 4).forEach((col) => pushUnique(col));
  if (chosen.length < 6) {
    available.forEach((col) => {
      if (chosen.length >= 8) return;
      pushUnique(col);
    });
  }
  return chosen.slice(0, 8);
}

function renderTable(rows, columns) {
  currentPreviewBaseRows = (rows || []).slice(0, 50);
  const availableColumns = Array.isArray(columns) ? columns : [];
  const selectedColumns = state.previewCompact ? getCompactPreviewColumns(availableColumns) : availableColumns;
  currentPreviewColumns = selectedColumns.length ? selectedColumns : availableColumns;
  currentTableRows = [...currentPreviewBaseRows];
  currentSort = { key: null, direction: "asc" };

  table.innerHTML = "";
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  (currentPreviewColumns || []).forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.addEventListener("click", () => sortTable(col));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  (currentTableRows || []).forEach((row) => {
    const tr = document.createElement("tr");
    (currentPreviewColumns || []).forEach((col) => {
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
  if (!column || !(currentPreviewColumns || []).includes(column)) return;
  const direction = currentSort.key === column && currentSort.direction === "asc" ? "desc" : "asc";
  currentSort = { key: column, direction };

  const sorted = [...currentPreviewBaseRows].sort((a, b) => {
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
  table.innerHTML = "";
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  (currentPreviewColumns || []).forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.addEventListener("click", () => sortTable(col));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  (currentTableRows || []).forEach((row) => {
    const tr = document.createElement("tr");
    (currentPreviewColumns || []).forEach((col) => {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
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
  const denomValues = getAnalysisRows().map((row) => parseNumber(row[denominator])).filter((value) => value !== null);
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

function aggregateByDate(rows, dateColumn, metric, bucket, dateRange, metricType, aggregation = state.aggregation || "sum") {
  const range = dateRange || { start: null, end: null };
  const totals = new Map();
  const counts = new Map();
  const distinct = new Map();
  (rows || []).forEach((row) => {
    const dateValue = parseDate(row[dateColumn]);
    if (!dateValue) return;
    if (range.start && dateValue < range.start) return;
    if (range.end && dateValue > range.end) return;
    const key = bucket === "week" ? getWeekStart(dateValue) : formatDateInput(dateValue);
    if (aggregation === "count") {
      totals.set(key, (totals.get(key) || 0) + 1);
      counts.set(key, (counts.get(key) || 0) + 1);
      return;
    }
    const rawValue = row[metric];
    if (aggregation === "count_distinct") {
      const set = distinct.get(key) || new Set();
      if (rawValue != null && String(rawValue).trim() !== "") set.add(String(rawValue));
      distinct.set(key, set);
      totals.set(key, set.size);
      counts.set(key, set.size);
      return;
    }
    const metricValue = parseNumber(rawValue);
    if (metricValue === null) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const labels = Array.from(totals.keys()).sort();
  const values = labels.map((label) => {
    const total = totals.get(label) || 0;
    const count = counts.get(label) || 1;
    if (aggregation === "avg" || metricType?.kind === "rate") return total / count;
    return total;
  });
  const ns = labels.map((label) => counts.get(label) || 0);
  return { labels, values, counts: ns };
}

function aggregateByCategory(rows, dimension, metric, topN, metricType, aggregation = state.aggregation || "sum") {
  if (!dimension) {
    return { labels: [], values: [], counts: [] };
  }
  const totals = new Map();
  const counts = new Map();
  const distinct = new Map();
  (rows || []).forEach((row) => {
    const key = row[dimension] ? String(row[dimension]).trim() : "Unknown";
    if (aggregation === "count") {
      totals.set(key, (totals.get(key) || 0) + 1);
      counts.set(key, (counts.get(key) || 0) + 1);
      return;
    }
    const rawValue = row[metric];
    if (aggregation === "count_distinct") {
      const set = distinct.get(key) || new Set();
      if (rawValue != null && String(rawValue).trim() !== "") set.add(String(rawValue));
      distinct.set(key, set);
      totals.set(key, set.size);
      counts.set(key, set.size);
      return;
    }
    const metricValue = parseNumber(rawValue);
    if (metricValue === null) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const entries = Array.from(totals.entries())
    .map(([label, total]) => ({
      label,
      value: (aggregation === "avg" || metricType?.kind === "rate") ? total / (counts.get(label) || 1) : total,
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

function buildTopCategoriesWithOther(rows, dimension, metric, metricType, topN = 8) {
  const totals = new Map();
  const counts = new Map();
  (rows || []).forEach((row) => {
    const key = row?.[dimension] ? String(row[dimension]).trim() : "Unknown";
    const metricValue = parseNumber(row?.[metric]);
    if (metricValue === null) return;
    totals.set(key, (totals.get(key) || 0) + metricValue);
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const sorted = Array.from(totals.entries())
    .map(([label, total]) => {
      const count = counts.get(label) || 0;
      return {
        label,
        total,
        count,
        value: metricType?.kind === "rate" ? (count ? total / count : 0) : total,
      };
    })
    .sort((a, b) => b.value - a.value);

  const head = sorted.slice(0, topN);
  const tail = sorted.slice(topN);
  if (tail.length) {
    const otherTotal = tail.reduce((sum, item) => sum + item.total, 0);
    const otherCount = tail.reduce((sum, item) => sum + item.count, 0);
    head.push({
      label: "Other",
      total: otherTotal,
      count: otherCount,
      value: metricType?.kind === "rate" ? (otherCount ? otherTotal / otherCount : 0) : otherTotal,
    });
  }

  return {
    labels: head.map((item) => item.label),
    values: head.map((item) => item.value),
    counts: head.map((item) => item.count),
  };
}

function truncateLabel(text, max = 22) {
  const value = String(text || "");
  if (value.length <= max) return value;
  return `${value.slice(0, Math.max(0, max - 1))}\u2026`;
}

function setChartCardEmptyState(canvasId, message = "") {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const card = canvas.closest(".chart-card");
  if (!card) return;
  const container = canvas.closest(".chartWrap") || card;
  let messageEl = container.querySelector(".card-empty-message");
  const text = String(message || "").trim();
  if (!text) {
    if (messageEl) messageEl.remove();
    canvas.classList.remove("hidden");
    return;
  }
  if (!messageEl) {
    messageEl = document.createElement("div");
    messageEl.className = "card-empty-message";
    container.appendChild(messageEl);
  }
  messageEl.textContent = text;
  canvas.classList.add("hidden");
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
      maintainAspectRatio: false,
      animation: false,
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
      maintainAspectRatio: false,
      animation: false,
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

function createHorizontalBarChart(canvasId, labels, values, metricType, counts = []) {
  const ctx = document.getElementById(canvasId);
  const fullLabels = labels.map((label) => String(label || ""));
  const displayLabels = fullLabels.map((label) => truncateLabel(label, 22));
  return new Chart(ctx, {
    type: "bar",
    data: {
      labels: displayLabels,
      datasets: [
        {
          label: "Total",
          data: values,
          backgroundColor: "#8b5cf6",
          borderRadius: 8,
          barThickness: 14,
          maxBarThickness: 16,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { right: 28 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: (items) => {
              const index = items?.[0]?.dataIndex ?? 0;
              return fullLabels[index] || displayLabels[index] || "";
            },
            label: (context) => {
              const n = counts?.[context.dataIndex] ?? null;
              const valueLabel = formatMetricValue(context.parsed.x, metricType);
              return n ? `${valueLabel} (n=${n})` : valueLabel;
            },
          },
        },
        datalabels: {
          color: "#ffffff",
          anchor: "end",
          align: "right",
          offset: 4,
          formatter: (value) => formatMetricValue(value, metricType),
          clamp: true,
          clip: false,
        },
      },
      scales: {
        y: {
          ticks: {
            autoSkip: false,
            maxRotation: 0,
            minRotation: 0,
          },
        },
        x: metricType?.kind === "rate"
          ? {
              min: 0,
              max: 1,
              ticks: {
                maxRotation: 0,
                minRotation: 0,
                callback: (value) => `${(Number(value) * 100).toFixed(0)}%`,
              },
            }
          : {
              beginAtZero: true,
              ticks: {
                maxRotation: 0,
                minRotation: 0,
              },
            },
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
      maintainAspectRatio: false,
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
    return new Chart(ctx, {
      type: "bar",
      data: { labels: [], datasets: [] },
      options: { responsive: true, maintainAspectRatio: false },
    });
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
  const metricQuality = classifyNumericMetricColumn(
    state.selections.primaryMetric,
    state.schema.profiles[state.selections.primaryMetric],
    rows.length
  );
  const warningTag = metricQuality.isIdentifierLike
    ? `<span class="metric-warning-tag">Identifier-like metric, insights may be low quality</span>`
    : "";
  const whereLiftBullets = (insight.whereLiftBullets || ["No strong subgroup contributors detected."])
    .map((item) => `<li>${escapeHtml(item)}</li>`)
    .join("");
  const subgroupBreakdownRows = (insight.subgroupBreakdown || []).slice(0, 12).map((row) => `
    <li>${escapeHtml(`${row.primarySegment} × ${row.secondaryDimension}:${row.secondaryValue} | ${row.deltaLabel} | ${percentFormatter.format(row.shareOfDelta || 0)} of subgroup lift`)}</li>
  `).join("");
  const subgroupMarkup = `
    <details class="insight-diagnostics">
      <summary>View subgroup breakdown</summary>
      <ul class="insight-mini-list">
        ${subgroupBreakdownRows || "<li>No subgroup breakdown available.</li>"}
      </ul>
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
    </section>
    <section class="insight-section">
      <h5>Impact</h5>
      <p class="insight-meta">${escapeHtml(insight.impactSummary || "Impact framing unavailable.")}</p>
    </section>
    <section class="insight-section">
      <h5>Where lift occurs</h5>
      <ul class="insight-mini-list">${whereLiftBullets}</ul>
    </section>
    <section class="insight-section">
      <h5>Stability classification</h5>
      <p class="insight-meta">${escapeHtml(insight.stabilitySummary || "Insufficient subgroup evidence to classify stability.")}</p>
    </section>
    <section class="insight-section">
      <h5>Recommended action</h5>
      <p class="insight-meta">${escapeHtml(insight.action || "Investigate key drivers and validate with a controlled test.")}</p>
    </section>
    ${subgroupMarkup}
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
    const hasDateField = Boolean(state.dateColumn && state.datasetCapabilities?.hasDateField);
    return hasDateField
      ? "Evidence is directionally useful, but trend or segment validation is limited."
      : "Evidence is directionally useful, but segment validation is limited.";
  }
  return "Evidence quality is moderate and should be validated with additional slices.";
}

function renderUserInsightEvidence(panel, insight) {
  if (!panel) return;
  openEvidenceDrawer(insight);
}

function openEvidenceDrawer(insight) {
  if (!evidenceDrawer || !evidenceContent) return;
  evidenceDrawer.classList.remove("hidden");
  evidenceDrawer.setAttribute("aria-hidden", "false");
  if (evidenceScrim) evidenceScrim.classList.remove("hidden");
  evidenceContent.innerHTML = buildEvidenceDrawerContent(insight);
}

function closeEvidenceDrawer() {
  if (!evidenceDrawer || !evidenceContent) return;
  evidenceDrawer.classList.add("hidden");
  evidenceDrawer.setAttribute("aria-hidden", "true");
  if (evidenceScrim) evidenceScrim.classList.add("hidden");
  evidenceContent.innerHTML = "";
}

function buildEvidenceCategorySummary(rows, dimension, metric, metricType, limit = 5) {
  if (!dimension || !metric) return { rows: [], denominator: null, totalContribution: 0 };
  const denominator = metricType?.kind === "rate" ? findRateDenominator(metric, state.schema.numeric || []) : null;
  const byCategory = new Map();
  (rows || []).forEach((row) => {
    const category = String(row?.[dimension] ?? "Unknown").trim() || "Unknown";
    const metricValue = parseNumber(row?.[metric]);
    if (metricValue === null) return;
    const weightRaw = denominator ? parseNumber(row?.[denominator]) : null;
    const weight = denominator ? Math.max(0, Number(weightRaw) || 0) : 1;
    const contribution = metricType?.kind === "rate"
      ? (denominator ? (metricValue * weight) : metricValue)
      : metricValue;
    const current = byCategory.get(category) || {
      category,
      sumMetric: 0,
      rowCount: 0,
      weightSum: 0,
      contribution: 0,
    };
    current.sumMetric += metricValue;
    current.rowCount += 1;
    current.weightSum += weight;
    current.contribution += contribution;
    byCategory.set(category, current);
  });

  const ranked = Array.from(byCategory.values())
    .map((entry) => ({
      category: entry.category,
      rowCount: entry.rowCount,
      contribution: entry.contribution,
      value: metricType?.kind === "rate"
        ? (denominator
            ? (entry.weightSum > 0 ? entry.contribution / entry.weightSum : 0)
            : (entry.rowCount > 0 ? entry.sumMetric / entry.rowCount : 0))
        : entry.contribution,
    }))
    .sort((a, b) => b.contribution - a.contribution);

  const head = ranked.slice(0, limit);
  const tail = ranked.slice(limit);
  if (tail.length) {
    const otherContribution = tail.reduce((sum, item) => sum + item.contribution, 0);
    const otherRows = tail.reduce((sum, item) => sum + item.rowCount, 0);
    const otherWeightedSum = tail.reduce((sum, item) => sum + item.value * item.rowCount, 0);
    head.push({
      category: "Other",
      rowCount: otherRows,
      contribution: otherContribution,
      value: metricType?.kind === "rate"
        ? (otherRows > 0 ? otherWeightedSum / otherRows : 0)
        : otherContribution,
    });
  }

  const totalContribution = head.reduce((sum, item) => sum + item.contribution, 0);
  const rowsWithMetric = head.reduce((sum, item) => sum + item.rowCount, 0);
  const rowsWithShare = head.map((item) => ({
    ...item,
    share: totalContribution > 0 ? item.contribution / totalContribution : 0,
  }));
  return { rows: rowsWithShare, denominator, totalContribution, rowsWithMetric };
}

function buildEvidenceDriverCorrelationTable(rows, metric, limit = 5) {
  const metrics = (state.schema.numeric || []).filter((candidate) => candidate && candidate !== metric);
  const ranked = metrics
    .map((candidate) => {
      let paired = 0;
      (rows || []).forEach((row) => {
        const x = parseNumber(row?.[metric]);
        const y = parseNumber(row?.[candidate]);
        if (x !== null && y !== null) paired += 1;
      });
      const corr = computeCorrelation(rows, metric, candidate);
      if (corr === null) return null;
      return { metric: candidate, corr, paired };
    })
    .filter(Boolean)
    .sort((a, b) => Math.abs(b.corr) - Math.abs(a.corr))
    .slice(0, limit);
  return ranked;
}

function describeEvidenceAggregation(metricType, metric, denominator) {
  if (metricType?.kind === "rate") {
    if (denominator) {
      return `Weighted average. Value = SUM(${formatMetricLabel(metric)} × ${formatMetricLabel(denominator)}) / SUM(${formatMetricLabel(denominator)}). Shares use weighted contribution.`;
    }
    return `Average of row-level rates. Value = AVG(${formatMetricLabel(metric)}). Shares use summed rates as a descriptive proxy.`;
  }
  return `Sum aggregation. Value = SUM(${formatMetricLabel(metric)}). Shares are based on each category's share of total summed value.`;
}

function buildEvidenceDrawerContent(insight) {
  const rows = getAnalysisRows();
  const metric = insight?.metricKey || state.selections.primaryMetric || "";
  const dimension = insight?.dimensionKey || state.selectedDimension || "";
  const metricValues = rows.map((row) => parseNumber(row?.[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);
  const evidence = insight?.evidence || {};
  const categorySummary = buildEvidenceCategorySummary(rows, dimension, metric, metricType, 5);
  const corrRankings = buildEvidenceDriverCorrelationTable(rows, metric, 5);
  const confidence = insight?.confidence || {};
  const confidenceReasons = (confidence?.reasons || []).slice(0, 3);
  const confidenceMetrics = confidence?.metrics || {};
  const topCategories = categorySummary.rows || [];
  const top = topCategories.find((item) => item.category !== "Other") || topCategories[0];
  const runnerUp = topCategories.filter((item) => item.category !== "Other")[1] || topCategories[1];
  const deltaRaw = top && runnerUp ? (top.value - runnerUp.value) : null;
  const deltaPct = top && runnerUp && Math.abs(runnerUp.value) > 0 ? (deltaRaw / runnerUp.value) : null;
  const confidenceRule = "Confidence starts high and is downgraded for no date context, fewer than 3 periods, low sample size (<8), high variance (>0.4), or a small delta (<10%).";
  const hasStatValidation = Boolean(evidence?.trendSummary?.periodsAnalysed >= 3 && (confidenceMetrics.sampleSize || 0) >= 8);
  const aggregationText = describeEvidenceAggregation(metricType, metric, categorySummary.denominator);
  const categoryRows = topCategories.map((item) => `
    <tr>
      <td>${escapeHtml(item.category)}</td>
      <td>${escapeHtml(formatMetricValue(item.value, metricType))}</td>
      <td>${escapeHtml(percentFormatter.format(item.share || 0))}</td>
      <td>${escapeHtml(numberFormatter.format(item.rowCount || 0))}</td>
    </tr>
  `).join("");
  const driverRows = corrRankings.map((item) => `
    <tr>
      <td>${escapeHtml(formatMetricLabel(item.metric))}</td>
      <td>${escapeHtml(item.corr.toFixed(3))}</td>
      <td>${item.corr >= 0 ? "Positive" : "Negative"}</td>
      <td>${escapeHtml(numberFormatter.format(item.paired || 0))}</td>
    </tr>
  `).join("");

  return `
    <div class="evidence-section">
      <strong>Insight being validated</strong>
      <p class="insight-meta">${escapeHtml(insight?.headline || insight?.title || "No insight statement available.")}</p>
    </div>
    <div class="evidence-section">
      <strong>Analytical method</strong>
      <div class="evidence-definition-grid">
        <div><strong>Primary metric</strong><div>${escapeHtml(formatMetricLabel(metric || "—"))}</div></div>
        <div><strong>Dimension</strong><div>${escapeHtml(formatMetricLabel(dimension || "—"))}</div></div>
        <div><strong>Aggregation logic</strong><div>${escapeHtml(aggregationText)}</div></div>
        <div><strong>Sample size</strong><div>${escapeHtml(numberFormatter.format(categorySummary.rowsWithMetric || metricValues.length || 0))} rows with metric values</div></div>
      </div>
      ${!hasStatValidation ? `<p class="evidence-callout">This is a descriptive comparison, not a causal test.</p>` : ""}
    </div>
    <div class="evidence-section">
      <strong>Category totals and share of total (Top 5 + Other)</strong>
      <div class="evidence-table-wrap">
        <table class="evidence-table">
          <thead><tr><th>Category</th><th>Value</th><th>Share of total</th><th>Rows</th></tr></thead>
          <tbody>${categoryRows || `<tr><td colspan="4">No category totals available.</td></tr>`}</tbody>
        </table>
      </div>
    </div>
    <div class="evidence-section">
      <strong>Delta calculation</strong>
      <div class="evidence-formula">
        ${top && runnerUp
          ? `Δ = ${escapeHtml(formatMetricLabel(metric))}(${escapeHtml(top.category)}) - ${escapeHtml(formatMetricLabel(metric))}(${escapeHtml(runnerUp.category)})`
          : "Delta unavailable: fewer than two comparable categories."}
      </div>
      ${top && runnerUp
        ? `<p class="helper-text">Result: ${escapeHtml(formatMetricValue(deltaRaw, metricType))}${deltaPct === null ? "" : ` (${escapeHtml((deltaPct * 100).toFixed(1))}% vs runner-up)`}.</p>`
        : ""}
    </div>
    <div class="evidence-section">
      <strong>Confidence logic</strong>
      <p class="helper-text">${escapeHtml(confidenceRule)}</p>
      <div class="evidence-definition-grid">
        <div><strong>Confidence level</strong><div>${escapeHtml(String(confidence.level || "medium").toUpperCase())}</div></div>
        <div><strong>Periods analyzed</strong><div>${escapeHtml(String(confidenceMetrics.periods ?? evidence?.trendSummary?.periodsAnalysed ?? 0))}</div></div>
        <div><strong>Sample size used</strong><div>${escapeHtml(String(confidenceMetrics.sampleSize ?? categorySummary.rowsWithMetric ?? 0))}</div></div>
        <div><strong>Variance score</strong><div>${escapeHtml(Number(confidenceMetrics.variance ?? evidence?.trendSummary?.variance ?? 0).toFixed(2))}</div></div>
      </div>
      <ul class="insight-mini-list">
        ${(confidenceReasons.length ? confidenceReasons : ["No additional confidence caveats provided."])
          .map((reason) => `<li>${escapeHtml(reason)}</li>`).join("")}
      </ul>
    </div>
    <div class="evidence-section">
      <strong>Driver ranking (correlation coefficients)</strong>
      <div class="evidence-table-wrap">
        <table class="evidence-table">
          <thead><tr><th>Driver metric</th><th>Correlation (r)</th><th>Direction</th><th>Paired rows</th></tr></thead>
          <tbody>${driverRows || `<tr><td colspan="4">No numeric driver correlations available.</td></tr>`}</tbody>
        </table>
      </div>
    </div>
  `;
}

function selectSecondaryDimensions(rows, primaryDimension, metric, metricType, limit = 2) {
  const candidates = (state.schema.categoricals || [])
    .filter((dim) => dim && dim !== primaryDimension)
    .filter((dim) => {
      const uniqueCount = state.schema.profiles?.[dim]?.uniqueCount || 0;
      return uniqueCount >= 2 && uniqueCount <= 12;
    });
  const scored = candidates.map((dim) => {
    const grouped = aggregateByCategory(rows, dim, metric, 12, metricType);
    const values = grouped.values || [];
    if (values.length < 2) return null;
    const mean = values.reduce((sum, value) => sum + (Number(value) || 0), 0) / values.length;
    const variance = values.reduce((sum, value) => sum + (((Number(value) || 0) - mean) ** 2), 0) / values.length;
    const std = Math.sqrt(variance);
    if (!Number.isFinite(std) || std <= 0) return null;
    return { dim, std, uniqueCount: state.schema.profiles?.[dim]?.uniqueCount || 0 };
  }).filter(Boolean);
  scored.sort((a, b) => (b.std - a.std) || (a.uniqueCount - b.uniqueCount));
  const picked = scored.slice(0, limit).map((item) => item.dim);
  if (picked.length) return picked;
  return candidates.slice(0, Math.max(1, limit));
}

function aggregateBySecondary(rows, secondaryDim, metric, metricType) {
  const map = new Map();
  (rows || []).forEach((row) => {
    const key = String(row?.[secondaryDim] ?? "Unknown").trim() || "Unknown";
    const metricValue = parseNumber(row?.[metric]);
    if (metricValue === null) return;
    const current = map.get(key) || { sum: 0, count: 0 };
    current.sum += metricValue;
    current.count += 1;
    map.set(key, current);
  });
  const out = new Map();
  map.forEach((value, key) => {
    out.set(key, metricType?.kind === "rate" ? (value.count ? value.sum / value.count : 0) : value.sum);
  });
  return out;
}

function buildCrossDimensionContributors({ rows, primaryDimension, winnerSegment, runnerUpSegment, secondaryDims, metric, metricType }) {
  const winnerRows = (rows || []).filter((row) => String(row?.[primaryDimension] ?? "Unknown").trim() === String(winnerSegment));
  const runnerRows = (rows || []).filter((row) => String(row?.[primaryDimension] ?? "Unknown").trim() === String(runnerUpSegment));
  const allContributors = [];
  (secondaryDims || []).forEach((secondaryDim) => {
    const winAgg = aggregateBySecondary(winnerRows, secondaryDim, metric, metricType);
    const runAgg = aggregateBySecondary(runnerRows, secondaryDim, metric, metricType);
    const keys = new Set([...winAgg.keys(), ...runAgg.keys()]);
    keys.forEach((key) => {
      const winnerValue = Number(winAgg.get(key) || 0);
      const runnerValue = Number(runAgg.get(key) || 0);
      const delta = winnerValue - runnerValue;
      if (!Number.isFinite(delta) || delta === 0) return;
      allContributors.push({
        primarySegment: winnerSegment,
        secondaryDimension: secondaryDim,
        secondaryValue: key,
        winnerValue,
        runnerValue,
        delta,
      });
    });
  });
  allContributors.sort((a, b) => b.delta - a.delta);
  const positiveTotal = allContributors.reduce((sum, item) => sum + Math.max(0, item.delta), 0);
  const ranked = (allContributors.filter((item) => item.delta > 0).length
    ? allContributors.filter((item) => item.delta > 0)
    : allContributors)
    .slice(0, 12)
    .map((item) => ({
      ...item,
      shareOfDelta: positiveTotal > 0 ? Math.max(0, item.delta) / positiveTotal : 0,
      deltaLabel: formatMetricValue(item.delta, metricType),
    }));
  const topThree = ranked.slice(0, 3);
  const topShare = topThree.length ? (topThree[0].shareOfDelta || 0) : 0;
  return { topThree, ranked, positiveTotal, topShare };
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
  const secondaryDims = selectSecondaryDimensions(rows, dimension, metric, metricType, 2);
  const cross = buildCrossDimensionContributors({
    rows,
    primaryDimension: dimension,
    winnerSegment: top.segment,
    runnerUpSegment: runnerUp.segment,
    secondaryDims,
    metric,
    metricType,
  });
  const winnerShare = totalValue > 0 ? (top.rawValue || 0) / totalValue : 0;
  const deltaOfTotal = totalValue > 0 ? (deltaRaw || 0) / totalValue : 0;
  const topContributors = (cross.ranked || []).slice(0, 2);
  const topContributor = topContributors[0] || null;
  const topContributorShare = topContributor?.shareOfDelta || 0;
  const topTwoShare = topContributors.reduce((sum, item) => sum + (item.shareOfDelta || 0), 0);
  const topContributorLabels = topContributors.map((entry) => `${formatMetricLabel(entry.secondaryDimension)}=${entry.secondaryValue}`);
  const contributorHeadline = topContributorLabels.length >= 2
    ? `${topContributorLabels[0]} and ${topContributorLabels[1]}`
    : (topContributorLabels[0] || "available subgroups");
  const whereLiftBullets = topContributors.map((entry) => {
    const shareText = percentFormatter.format(entry.shareOfDelta || 0);
    return `${formatMetricLabel(entry.secondaryDimension)}=${entry.secondaryValue} contributes ${entry.deltaLabel} (${shareText} of delta).`;
  });
  if (topContributors.length >= 2) {
    whereLiftBullets.push(`Top 2 contributors explain ${percentFormatter.format(topTwoShare)} of total delta.`);
  } else if (topContributors.length === 1) {
    whereLiftBullets.push(`Top contributor explains ${percentFormatter.format(topTwoShare)} of total delta.`);
  } else {
    whereLiftBullets.push("No subgroup contributor stands out in this cut.");
  }
  let stabilityLabel = "Broadly distributed";
  if (topContributorShare > 0.5) stabilityLabel = "Highly concentrated";
  else if (topContributorShare >= 0.3) stabilityLabel = "Moderately concentrated";
  const stabilitySummary = `${stabilityLabel}: top subgroup contributes ${percentFormatter.format(topContributorShare)} of delta.`;
  const winnerTag = `${formatMetricLabel(dimension)}=${top.segment}`;
  let action = `For ${winnerTag}: Scale winning dimension broadly while validating marginal ROI by subgroup.`;
  if (stabilityLabel === "Highly concentrated") {
    action = `For ${winnerTag}: Focus expansion on dominant subgroup while diversifying to reduce risk.`;
  } else if (stabilityLabel === "Moderately concentrated") {
    action = `For ${winnerTag}: Prioritize top cohorts and test incremental lift before scaling.`;
  }
  const periodSplit = splitRowsIntoCurrentPrior(rows);
  const confidence = computeUserUnifiedConfidence({
    periods: periodSplit.hasDateField ? periodSplit.periodsAnalysed : 0,
    sampleSize: comparison.counts?.[0] || 0,
    deltaPercent: Math.abs(deltaPercent || 0),
    variance: 0,
    hasDateField: periodSplit.hasDateField,
  });

  const insight = {
    id: `user:${metric}:${dimension}:${top.segment}`,
    headline: `${top.segment} drives +${formatMetricValue(deltaRaw, metricType)} vs ${runnerUp.segment}, led by ${contributorHeadline}.`,
    deltaSummary,
    metricKey: metric,
    dimensionKey: dimension,
    topSegment: top.segment,
    secondaryDimensions: secondaryDims,
    whereLiftBullets,
    subgroupBreakdown: cross.ranked || [],
    stabilitySummary,
    stabilityLabel,
    confidence,
    action,
    impactSummary: `Winner share ${percentFormatter.format(winnerShare)} of total ${formatMetricLabel(metric)}. Delta equals ${percentFormatter.format(deltaOfTotal)} of total.`,
    evidence: {
      comparisonTable: comparisonTable.slice(0, 8),
      contributionSummary: whereLiftBullets[0] || "No subgroup contribution summary available.",
      trendSummary: periodSplit.hasDateField ? {
        periodsAnalysed: periodSplit.periodsAnalysed,
        variance: "—",
      } : null,
      consistencyScore: null,
      driverBreakdownTable: (cross.ranked || []).slice(0, 5).map((row) => ({
        combination: `${formatMetricLabel(row.secondaryDimension)}=${row.secondaryValue}`,
        current: formatMetricValue(row.winnerValue, metricType),
        prior: formatMetricValue(row.runnerValue, metricType),
        delta: row.deltaLabel,
        liftShare: percentFormatter.format(row.shareOfDelta || 0),
      })),
      driverNarrative: whereLiftBullets[0] || "No dominant subgroup narrative available.",
      filtersUsed: [
        `Metric: ${formatMetricLabel(metric)}`,
        `Primary dimension: ${dimension}`,
        `Secondary dimensions: ${(secondaryDims || []).map((dim) => formatMetricLabel(dim)).join(", ") || "none"}`,
        periodSplit.hasDateField ? "Window: Last 30 days vs prior 30 days" : "Window: Snapshot (no date field)",
      ],
    },
    executiveSummary: `${top.segment} leads ${runnerUp.segment} by ${formatMetricValue(deltaRaw, metricType)} on ${formatMetricLabel(metric)}.`,
    diagnostics: {
      anomalies: [],
      concentrationRisk: [stabilitySummary],
      relationships: whereLiftBullets.slice(0, 2),
    },
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
      hasDate ? "Not enough reliable metric semantics for trend-based claims." : "No historical coverage available.",
    ],
    metrics: { periods: hasDate ? 1 : 0, sampleSize: rows.length, deltaPercent: 0, variance: 0 },
  };
  return {
    id: `safe:${metric}:${dimension}`,
    headline: "Structured tables detected, but insight quality is limited.",
    title: "Structured tables detected, but insight quality is limited.",
    executiveSummary: hasDate
      ? "This file contains structured tables but metric semantics are limited. Review top segments by spend and conversions where available."
      : "This file contains structured tables but metric semantics are limited. Review top segments by spend and conversions where available.",
    executiveBullets: [
      `Use ${dimension} to review top segments by core marketing metrics.`,
      "Prefer Spend, Clicks, Impressions, Conversions, CTR, CPC, CPA, or eCPM as the primary metric.",
    ],
    deltaSummary: "Descriptive mode only; change claims are suppressed.",
    whereLiftBullets: ["No subgroup lift analysis available because the selected metric is not suitable."],
    subgroupBreakdown: [],
    stabilitySummary: "Stability classification unavailable in descriptive mode.",
    stabilityLabel: "Unavailable",
    impactSummary: "Impact framing unavailable in descriptive mode.",
    action: "Select a marketing metric such as Spend, Conversions, Clicks, or Impressions to generate higher-confidence insights.",
    confidence,
    diagnostics: {
      anomalies: ["Change-based insights are blocked until a meaningful metric is selected."],
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
        hasDate ? "Date context detected but not used for claims" : "No date field detected",
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
  const categoryValues = (comparisonTable || [])
    .map((item) => ({
      segment: item?.segment || "Unknown",
      value: Number(item?.rawValue),
    }))
    .filter((item) => Number.isFinite(item.value));
  if (categoryValues.length >= 3) {
    const mean = categoryValues.reduce((sum, item) => sum + item.value, 0) / categoryValues.length;
    const variance = categoryValues.reduce((sum, item) => sum + ((item.value - mean) ** 2), 0) / categoryValues.length;
    const std = Math.sqrt(variance);
    if (std > 0) {
      const zScoreAnomalies = categoryValues
        .map((item) => {
          const deviation = item.value - mean;
          const z = deviation / std;
          return { ...item, z, deviation };
        })
        .filter((item) => Math.abs(item.z) >= 2)
        .sort((a, b) => Math.abs(b.z) - Math.abs(a.z));
      zScoreAnomalies.forEach((item) => {
        const sign = item.deviation >= 0 ? "+" : "-";
        const deviationAbs = Math.abs(item.deviation);
        anomalies.push(
          `${item.segment}: z-score ${item.z.toFixed(2)}, deviation ${sign}${formatMetricValue(deviationAbs, metricType)} from mean.`
        );
      });
    }
  }
  const topShare = comparisonTable?.[0]?.rawValue && comparisonTable.length
    ? (comparisonTable[0].rawValue / Math.max(1, comparisonTable.reduce((s, r) => s + (r.rawValue || 0), 0)))
    : 0;
  const totalForShare = Math.max(1, (comparisonTable || []).reduce((s, r) => s + (r.rawValue || 0), 0));
  const topTwoShare = ((comparisonTable?.[0]?.rawValue || 0) + (comparisonTable?.[1]?.rawValue || 0)) / totalForShare;
  let concentrationLevel = "Low";
  let concentrationInterpretation = "Distribution is relatively balanced across categories.";
  if (topShare > 0.6) {
    concentrationLevel = "High";
    concentrationInterpretation = "Performance is highly concentrated in a single category and may be brittle.";
  } else if (topTwoShare > 0.8) {
    concentrationLevel = "Moderate";
    concentrationInterpretation = "Most performance is concentrated in two categories; diversification is limited.";
  }
  concentrationRisk.push(
    `Top category share: ${percentFormatter.format(topShare)} | Top 2 combined share: ${percentFormatter.format(topTwoShare)} | Risk level: ${concentrationLevel}. ${concentrationInterpretation}`
  );
  if (Math.abs(deltaPercent || 0) < 10) anomalies.push("Performance gap is small; treat the insight as directional.");
  if (state.dateColumn && driverInfo?.varianceScore != null && driverInfo.varianceScore > 0.4) {
    anomalies.push("Driver pattern is volatile across observed periods.");
  }
  const corrCandidates = (state.schema.numeric || []).filter((m) => m !== metric);
  let best = null;
  corrCandidates.forEach((candidate) => {
    const corr = computeCorrelation(rows, metric, candidate);
    if (corr === null) return;
    let pairedSampleSize = 0;
    (rows || []).forEach((row) => {
      const x = parseNumber(row?.[metric]);
      const y = parseNumber(row?.[candidate]);
      if (x !== null && y !== null) pairedSampleSize += 1;
    });
    if (!best || Math.abs(corr) > Math.abs(best.corr)) best = { candidate, corr, pairedSampleSize };
  });
  if (best) {
    const direction = best.corr >= 0 ? "positive" : "negative";
    const collinearityNote = Math.abs(best.corr) > 0.9 ? " High collinearity possible." : "";
    relationships.push(
      `${formatMetricLabel(best.candidate)} has the strongest relationship to ${formatMetricLabel(metric)} (r=${best.corr.toFixed(2)}, sample size=${best.pairedSampleSize}, direction=${direction}). Correlation indicates association, not causation.${collinearityNote}`
    );
  }
  if (!relationships.length) {
    relationships.push("No strong secondary metric relationships were detected. Correlation indicates association, not causation.");
  }
  if (!anomalies.length) anomalies.push("No statistically significant outliers detected at |z| ≥ 2.");
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
    reasons.push("No historical coverage available for this metric.");
  }
  if (hasDateField && periods < 3) {
    level = "low";
    reasons.push("Not enough periods to confirm a stable pattern.");
  }
  if (sampleSize < 8 && level !== "low") {
    level = "medium";
    reasons.push("Limited sample size in the top driver combination.");
  }
  if (hasDateField && (variance || 0) > 0.4 && level !== "low") {
    level = "medium";
    reasons.push("Metric pattern is volatile across periods.");
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

function buildUserDynamicAction({ metricKey, topSegment, driverCombination, confidence, deltaStrength, hasDateField }) {
  const name = formatMetricLabel(metricKey);
  if (confidence?.level === "low") {
    return hasDateField
      ? `Validate whether ${driverCombination} remains the primary driver with additional periods before reallocating budget.`
      : `Validate whether ${driverCombination} remains the primary driver with an additional segment review before reallocating budget.`;
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
  const hasDateField = Boolean(state.dateColumn && state.datasetCapabilities?.hasDateField);
  actions.push(`Investigate ${preferredDim} segments where ${metric} is lowest and compare Sessions/ActiveUsers to spot friction.`);
  actions.push(hasDateField
    ? `Run a cohort review for ${preferredDim} segments with low ${metric} and validate against ActiveUsers patterns over time.`
    : `Run a cohort review for ${preferredDim} segments with low ${metric} and validate against ActiveUsers patterns.`);
  const driverMetric = best?.metric || "Sessions";
  actions.push(`Quantify how changes in ${driverMetric} shift ${metric} and test a targeted uplift plan.`);

  insights.push({
    title: "Action items",
    summary: "Three next steps tied to observed segments/drivers.",
    bullets: actions.slice(0, 3).map((action) => `Next step: ${action}`),
  });

  return insights;
}

function renderDiagnosticsCorrelation(rows) {
  if (!diagCorrXSelect || !diagCorrYSelect || !diagCorrSummary || !diagCorrChartCanvas) return;
  const numericMetrics = Array.isArray(state.schema.numeric) ? state.schema.numeric : [];
  if (diagCorrCaution) diagCorrCaution.textContent = "";
  if (diagCorrExplain) diagCorrExplain.textContent = "";
  if (diagCorrTakeaway) diagCorrTakeaway.textContent = "";
  if (diagCorrTitle) diagCorrTitle.textContent = "Correlation analysis";
  if (diagCorrSubtitle) diagCorrSubtitle.textContent = "r=— • n=—";
  if (diagnosticsCorrelationChartInstance) {
    diagnosticsCorrelationChartInstance.destroy();
    diagnosticsCorrelationChartInstance = null;
  }
  if (numericMetrics.length < 2) {
    diagCorrSummary.textContent = "Correlation unavailable. At least 2 numeric metrics are required.";
    return;
  }

  const previousX = state.diagnosticsCorrelation?.xMetric;
  const previousY = state.diagnosticsCorrelation?.yMetric;
  const defaultX = numericMetrics.includes(previousX) ? previousX : (state.selections.primaryMetric || numericMetrics[0]);
  const defaultYCandidate = numericMetrics.includes(previousY) ? previousY : numericMetrics.find((metric) => metric !== defaultX);
  const defaultY = defaultYCandidate && defaultYCandidate !== defaultX
    ? defaultYCandidate
    : numericMetrics.find((metric) => metric !== defaultX);

  diagCorrXSelect.innerHTML = "";
  diagCorrYSelect.innerHTML = "";
  numericMetrics.forEach((metric) => {
    const xOption = document.createElement("option");
    xOption.value = metric;
    xOption.textContent = formatMetricLabel(metric);
    diagCorrXSelect.appendChild(xOption);

    const yOption = document.createElement("option");
    yOption.value = metric;
    yOption.textContent = formatMetricLabel(metric);
    diagCorrYSelect.appendChild(yOption);
  });

  diagCorrXSelect.value = defaultX;
  let yMetric = defaultY || numericMetrics.find((metric) => metric !== defaultX);
  if (!yMetric) {
    diagCorrSummary.textContent = "Correlation unavailable. No valid metric pair found.";
    return;
  }
  if (yMetric === defaultX) {
    yMetric = numericMetrics.find((metric) => metric !== defaultX) || null;
  }
  diagCorrYSelect.value = yMetric;
  state.diagnosticsCorrelation = { xMetric: defaultX, yMetric };

  if (diagCorrXSelect.value === diagCorrYSelect.value) {
    const replacement = numericMetrics.find((metric) => metric !== diagCorrXSelect.value);
    if (!replacement) {
      diagCorrSummary.textContent = "Correlation unavailable. Pick two different metrics.";
      return;
    }
    diagCorrYSelect.value = replacement;
    state.diagnosticsCorrelation.yMetric = replacement;
  }

  const xMetric = diagCorrXSelect.value;
  const yMetricSelected = diagCorrYSelect.value;
  const points = rows
    .map((row) => {
      const xVal = parseNumber(row?.[xMetric]);
      const yVal = parseNumber(row?.[yMetricSelected]);
      if (xVal === null || yVal === null) return null;
      return { x: xVal, y: yVal };
    })
    .filter(Boolean)
    .slice(0, 1200);
  if (points.length < 3) {
    diagCorrSummary.textContent = "Correlation unavailable. Not enough paired rows.";
    return;
  }

  const correlation = computeCorrelation(rows, xMetric, yMetricSelected);
  const regression = computeLinearRegression(points);
  if (correlation === null || !regression) {
    diagCorrSummary.textContent = "Correlation unavailable. Metric pair has zero variance.";
    return;
  }
  const absR = Math.abs(correlation);
  const direction = correlation >= 0 ? "positive" : "negative";
  const strengthLabel = absR >= 0.7 ? "Strong" : absR >= 0.3 ? "Moderate" : "Weak";
  const xLower = String(xMetric || "").toLowerCase();
  const yLower = String(yMetricSelected || "").toLowerCase();
  const xIsSpend = /spend|cost|budget/.test(xLower);
  const yIsSpend = /spend|cost|budget/.test(yLower);
  const xIsImpressions = /impression/.test(xLower);
  const yIsImpressions = /impression/.test(yLower);
  const xIsOutcome = /conversion|revenue|sales|gmv|arr|mrr/.test(xLower);
  const yIsOutcome = /conversion|revenue|sales|gmv|arr|mrr/.test(yLower);
  let actionLine = "Use this as directional evidence, then validate with segment-level tests before making budget changes.";
  if (correlation >= 0.7 && ((xIsSpend && yIsImpressions) || (yIsSpend && xIsImpressions))) {
    actionLine = "This suggests delivery scales with investment. Validate efficiency by checking eCPM trends and conversion outcomes.";
  } else if (correlation >= 0.7 && ((xIsSpend && yIsOutcome) || (yIsSpend && xIsOutcome))) {
    actionLine = "This suggests spend aligns with outcomes. Validate marginal returns by segmenting by audience and comparing CPA or ROAS.";
  } else if (absR < 0.3) {
    actionLine = "Spend changes do not explain this outcome. Investigate targeting, creative, or inventory mix as primary drivers.";
  }
  if (diagCorrTitle) diagCorrTitle.textContent = `Correlation: ${formatMetricLabel(xMetric)} vs ${formatMetricLabel(yMetricSelected)}`;
  if (diagCorrSubtitle) diagCorrSubtitle.textContent = `r=${correlation.toFixed(3)} • n=${points.length}`;
  if (diagCorrExplain) {
    diagCorrExplain.innerHTML = `
      <strong>Meaning:</strong> Correlation shows how tightly two metrics move together. It does not prove causation.<br>
      <strong>Strength:</strong> ${strengthLabel} ${direction} association (r=${correlation.toFixed(3)}).
    `;
  }
  if (diagCorrTakeaway) {
    diagCorrTakeaway.textContent = `Action: ${actionLine}`;
  }
  if (diagCorrCaution && points.length < 30) {
    diagCorrCaution.textContent = "Low sample size. Interpret with caution.";
  }
  diagCorrSummary.textContent = `y = ${regression.slope.toFixed(3)}x + ${regression.intercept.toFixed(3)}`;
  diagnosticsCorrelationChartInstance = createDiagnosticsCorrelationChart("diagCorrChart", points, regression, xMetric, yMetricSelected);
  bindDiagnosticsCorrelationResizeObserver();
  forceDiagnosticsCorrelationResize(0);
}

function computeLinearRegression(points) {
  if (!Array.isArray(points) || points.length < 2) return null;
  const meanX = points.reduce((sum, point) => sum + point.x, 0) / points.length;
  const meanY = points.reduce((sum, point) => sum + point.y, 0) / points.length;
  let numerator = 0;
  let denominator = 0;
  points.forEach((point) => {
    const dx = point.x - meanX;
    numerator += dx * (point.y - meanY);
    denominator += dx * dx;
  });
  if (!denominator) return null;
  const slope = numerator / denominator;
  const intercept = meanY - (slope * meanX);
  return { slope, intercept };
}

function scheduleDiagnosticsCorrelationResize() {
  if (!diagnosticsCorrelationChartInstance) return;
  requestAnimationFrame(() => {
    if (!diagnosticsCorrelationChartInstance) return;
    diagnosticsCorrelationChartInstance.resize();
    diagnosticsCorrelationChartInstance.update("none");
  });
}

function forceDiagnosticsCorrelationResize(delayMs = 0) {
  if (!diagnosticsCorrelationChartInstance) return;
  setTimeout(() => {
    if (!diagnosticsCorrelationChartInstance) return;
    diagnosticsCorrelationChartInstance.resize();
    diagnosticsCorrelationChartInstance.update("none");
  }, delayMs);
}

function bindDiagnosticsCorrelationResizeObserver() {
  if (typeof ResizeObserver === "undefined") return;
  const wrap = document.querySelector(".corrChartWrap");
  if (!wrap) return;
  if (diagnosticsCorrelationResizeObserver) {
    diagnosticsCorrelationResizeObserver.disconnect();
    diagnosticsCorrelationResizeObserver = null;
  }
  diagnosticsCorrelationResizeObserver = new ResizeObserver(() => {
    scheduleDiagnosticsCorrelationResize();
  });
  diagnosticsCorrelationResizeObserver.observe(wrap);
}

function createDiagnosticsCorrelationChart(canvasId, points, regression, xMetric, yMetric) {
  const ctx = document.getElementById(canvasId);
  const xValues = points.map((point) => point.x);
  const minX = Math.min(...xValues);
  const maxX = Math.max(...xValues);
  const lineData = [
    { x: minX, y: regression.slope * minX + regression.intercept },
    { x: maxX, y: regression.slope * maxX + regression.intercept },
  ];
  const xType = inferMetricType(xMetric, xValues);
  const yType = inferMetricType(yMetric, points.map((point) => point.y));
  return new Chart(ctx, {
    type: "scatter",
    data: {
      datasets: [
        {
          label: "Regression",
          type: "line",
          data: lineData,
          borderColor: "#fbbf24",
          borderWidth: 2,
          pointRadius: 0,
          tension: 0,
          order: 1,
        },
        {
          label: "Observed",
          data: points,
          pointRadius: 3,
          pointHoverRadius: 4,
          backgroundColor: "rgba(124, 58, 237, 0.65)",
          borderColor: "rgba(226, 232, 240, 0.85)",
          borderWidth: 0.6,
          order: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: { display: true, position: "top", labels: { color: "#cbd5e1" } },
        tooltip: {
          callbacks: {
            label: (context) => {
              if (context.datasetIndex === 1) return "Regression line";
              return `${formatMetricLabel(xMetric)}: ${formatMetricValue(context.parsed.x, xType)} | ${formatMetricLabel(yMetric)}: ${formatMetricValue(context.parsed.y, yType)}`;
            },
          },
        },
      },
      scales: {
        x: { ticks: { color: "#94a3b8" }, grid: { color: "rgba(148, 163, 184, 0.15)" } },
        y: { ticks: { color: "#94a3b8" }, grid: { color: "rgba(148, 163, 184, 0.15)" } },
      },
    },
  });
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
    suggestedTrends.innerHTML = "<li>Load a sample to see suggested views.</li>";
    return;
  }
  const suggestions = [];
  if (state.dateColumn && state.schema.numeric.length) {
    const metric = state.selections.primaryMetric || state.schema.numeric[0];
    suggestions.push(`${metric} over ${state.dateColumn}`);
    if (state.schema.categoricals.length) {
      suggestions.push(`${metric} by ${state.schema.categoricals[0]}`);
    }
  } else if (state.schema.numeric.length && state.schema.categoricals.length) {
    suggestions.push(`${state.selections.primaryMetric || state.schema.numeric[0]} by ${state.schema.categoricals[0]}`);
  }

  if (!suggestions.length) {
    suggestions.push("Add a date or categorical column to unlock more suggestions.");
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

function exportCsvData() {
  const rows = getAnalysisRows();
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
  const rows = getAnalysisRows();
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
