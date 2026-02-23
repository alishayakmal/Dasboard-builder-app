const fileInput = document.getElementById("fileInput");
const fileInputInline = document.getElementById("fileInputInline");
const loadSample = document.getElementById("loadSample");
const loadSampleInline = document.getElementById("loadSampleInline");
const landingView = document.getElementById("landingView");
const appView = document.getElementById("appView");
const signInButton = document.getElementById("signInBtn");
const signOutButton = document.getElementById("signOutButton");
const startFreeButton = document.getElementById("startFreeBtn");
const signupModal = document.getElementById("signupModal");
const closeModalButton = document.getElementById("closeModal");
const signupForm = document.getElementById("signupForm");
const formError = document.getElementById("formError");
const leadNameInput = document.getElementById("signupName");
const leadEmailInput = document.getElementById("signupEmail");
const leadCompanyInput = document.getElementById("signupCompany");
const leadUseCaseInput = document.getElementById("signupUseCase");
const signupSubmitButton = document.getElementById("signupSubmit");
const sheetsUrlInput = document.getElementById("sheetsUrl");
const sheetsRangeInput = document.getElementById("sheetsRange");
const sheetsHelper = document.getElementById("sheetsHelper");
const loadSheetButton = document.getElementById("loadSheet");
const connectGoogleButton = document.getElementById("connectGoogle");
const disconnectGoogleButton = document.getElementById("disconnectGoogle");
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
const dashboard = document.getElementById("dashboard");
const datasetSummary = document.getElementById("datasetSummary");
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

const samplePath = "data-sample.csv";
// Use the deployed Apps Script Web App URL ending in /exec (not the editor URL).
const WEBHOOK_URL = "PASTE_APPS_SCRIPT_EXEC_URL_HERE";
// TODO: Replace with your Google OAuth Client ID for GitHub Pages.
const GOOGLE_CLIENT_ID = "PASTE_YOUR_CLIENT_ID.apps.googleusercontent.com";
let googleAccessToken = null;
let googleTokenClient = null;
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
};

document.addEventListener("DOMContentLoaded", init);

function init() {
  console.log("init started");
  if (buildStamp) {
    buildStamp.textContent = `Build: ${new Date().toISOString()}`;
  }

  window.addEventListener("hashchange", handleRoute);

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
    testApiButton.addEventListener("click", testApiConnection);
  }

  if (startFreeButton) startFreeButton.addEventListener("click", openModal);
  if (closeModalButton) closeModalButton.addEventListener("click", closeModal);
  if (signupModal) signupModal.addEventListener("click", (event) => {
    if (event.target === signupModal) closeModal();
  });
  if (signupForm) signupForm.addEventListener("submit", handleSignupSubmit);
  if (signInButton) signInButton.addEventListener("click", handleSignInClick);
  if (signOutButton) signOutButton.addEventListener("click", handleSignOutClick);
  if (loadSheetButton) loadSheetButton.addEventListener("click", loadSheetData);
  if (connectGoogleButton) connectGoogleButton.addEventListener("click", handleConnectGoogleClick);
  if (disconnectGoogleButton) disconnectGoogleButton.addEventListener("click", handleDisconnectGoogleClick);
  if (exportSheetsButton) exportSheetsButton.addEventListener("click", handleExportToSheets);
  if (downloadPdfButton) downloadPdfButton.addEventListener("click", () => window.print());
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
  updateGoogleStatus();
  checkWebhookHealth();
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

function getSignedIn() {
  return localStorage.getItem("signedIn") === "true";
}

function setSignedIn(value) {
  localStorage.setItem("signedIn", value ? "true" : "false");
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

function initGoogleAuth() {
  if (!window.google || !window.google.accounts || !window.google.accounts.oauth2) return null;
  if (googleTokenClient) return googleTokenClient;

  googleTokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly",
    callback: (response) => {
      if (response?.access_token) {
        googleAccessToken = response.access_token;
        updateGoogleStatus();
        showToast("Google connected");
      }
    },
  });

  return googleTokenClient;
}

function handleConnectGoogleClick() {
  const tokenClient = initGoogleAuth();
  if (!tokenClient) {
    showError("Google auth not available. Ensure Google Identity Services is loaded.");
    return;
  }
  tokenClient.requestAccessToken({ prompt: "consent" });
}

function handleDisconnectGoogleClick() {
  if (!googleAccessToken) return;
  const token = googleAccessToken;
  googleAccessToken = null;
  updateGoogleStatus();
  showToast("Google disconnected");
  if (window.google?.accounts?.oauth2?.revoke) {
    window.google.accounts.oauth2.revoke(token, () => {});
  }
}

function updateGoogleStatus() {
  if (googleStatus) {
    googleStatus.textContent = googleAccessToken ? "Connected" : "Not connected";
  }
  if (disconnectGoogleButton) {
    disconnectGoogleButton.disabled = !googleAccessToken;
  }
  if (connectGoogleButton) {
    connectGoogleButton.textContent = googleAccessToken ? "Reconnect Google" : "Connect Google";
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
  openModal();
}

function handleSignOutClick() {
  setSignedIn(false);
  localStorage.removeItem("currentUser");
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
}

function closeModal() {
  if (!signupModal) return;
  signupModal.classList.add("hidden");
  signupModal.setAttribute("aria-hidden", "true");
}

function handleSignupSubmit(event) {
  event.preventDefault();
  const name = leadNameInput?.value?.trim() || "";
  const email = leadEmailInput?.value?.trim() || "";
  const company = leadCompanyInput?.value?.trim() || "";
  const useCase = leadUseCaseInput?.value?.trim() || "";

  if (!name || !email) {
    showFormError("Please provide your name and email.");
    return;
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    showFormError("Please enter a valid email address.");
    return;
  }

  const lead = {
    name,
    email,
    company,
    useCase,
    source: "shalytics landing",
    userAgent: navigator.userAgent,
    createdAt: new Date().toISOString(),
  };

  saveLead(lead);
  setSignedIn(true);
  localStorage.setItem("currentUser", email);
  updateSignInButton();
  closeModal();
  showToast("Signed up");
  routeTo("app");
  logSignupEvent(lead);
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

function logSignupEvent(payload) {
  if (signupSubmitButton) signupSubmitButton.disabled = true;
  if (webhookStatus) webhookStatus.textContent = "Saving...";
  postWebhookPlain(payload)
    .then((result) => {
      if (result.ok) {
        if (webhookStatus) webhookStatus.textContent = "Saved to Google Sheets";
        showToast("Saved");
      } else {
        if (webhookStatus) webhookStatus.textContent = "Saved locally only";
        showToast("Saved locally");
      }
    })
    .catch(() => {
      if (webhookStatus) webhookStatus.textContent = "Saved locally only";
      showToast("Saved locally");
    })
    .finally(() => {
      if (signupSubmitButton) signupSubmitButton.disabled = false;
    });
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

async function logEvent(payload) {
  if (!WEBHOOK_URL) return { ok: true, status: 0, body: "" };
  try {
    const response = await fetch(WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      redirect: "follow",
    });
    const body = await response.text();
    console.log("Event log response", response.status, body);
    if (response.ok) return { ok: true, status: response.status, body };
    // If not ok, still try a no-cors fallback to ensure write happens.
    await fetch(WEBHOOK_URL, {
      method: "POST",
      mode: "no-cors",
      body: JSON.stringify(payload),
    });
    return { ok: false, status: response.status, body };
  } catch (error) {
    console.log("Event log error", error);
    try {
      await fetch(WEBHOOK_URL, {
        method: "POST",
        mode: "no-cors",
        body: JSON.stringify(payload),
      });
      return { ok: true, status: 0, body: "no-cors fallback" };
    } catch (fallbackError) {
      console.log("Event log fallback error", fallbackError);
      return { ok: false, status: 0, body: String(error) };
    }
  }
}

async function postWebhookPlain(payload) {
  if (!WEBHOOK_URL) return { ok: true, status: 0, body: "" };
  try {
    const response = await fetch(WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    if (response.type === "opaque") {
      console.log("Signup webhook response opaque");
      return { ok: true, status: 0, body: "opaque" };
    }
    const body = await response.text();
    const okText = body.trim().toLowerCase() === "ok";
    console.log("Signup webhook response", response.status, body);
    return { ok: response.ok && okText, status: response.status, body };
  } catch (error) {
    console.log("Signup webhook error", error);
    return { ok: false, status: 0, body: String(error) };
  }
}

function checkWebhookHealth() {
  if (!WEBHOOK_URL || !webhookStatus) return;
  if (WEBHOOK_URL.includes("PASTE_")) {
    webhookStatus.textContent = "Webhook not configured. Paste your /exec URL.";
    return;
  }
  fetch(WEBHOOK_URL)
    .then((response) => {
      if (response.ok) {
        webhookStatus.textContent = "Webhook online";
      } else {
        webhookStatus.textContent = "Webhook offline. Verify the Apps Script Web App URL and deployment.";
      }
    })
    .catch(() => {
      webhookStatus.textContent = "Webhook offline. Verify the Apps Script Web App URL and deployment.";
    });
}

function handleExportToSheets() {
  const user = localStorage.getItem("currentUser") || "anonymous";
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
  return match ? match[1] : "";
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
  const url = normalizeSheetsUrl(input);
  const sheetId = extractSheetId(input);
  const range = sheetsRangeInput?.value?.trim() || "A1:Z1000";
  if (!url) {
    showError("Please paste a Google Sheets CSV link.");
    return;
  }
  if (sheetsHelper) {
    sheetsHelper.textContent = "Paste a Google Sheets link. If it is an edit link, the app will convert it to a CSV export link.";
  }
  clearMessages();
  setStatus("Loading Google Sheet");
  try {
    if (sheetId) {
      if (!googleAccessToken) {
        showError("Connect Google to load private sheets.");
        return;
      }
      const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(sheetId)}/values/${encodeURIComponent(range)}?majorDimension=ROWS`;
      const privateResult = await fetch(apiUrl, {
        headers: { Authorization: `Bearer ${googleAccessToken}` },
      });
      if (privateResult.ok) {
        const payload = await privateResult.json();
        const csvText = valuesToCsv(payload.values || []);
        parseCsvText(csvText);
        switchTab("upload");
        return;
      }
      if (privateResult.status === 401 || privateResult.status === 403) {
        showError("Permission denied or token expired. Reconnect Google.");
        return;
      }
      const errorText = await privateResult.text();
      showError("Google Sheets API error.", `Details: ${errorText}`);
      return;
    }

    const text = await fetchSheetText(url);
    parseCsvText(text);
    switchTab("upload");
  } catch (error) {
    const fallback = buildSheetsFallback(url);
    if (fallback) {
      try {
        const text = await fetchSheetText(fallback);
        parseCsvText(text);
        switchTab("upload");
        return;
      } catch (fallbackError) {
        error = fallbackError;
      }
    }
    if (sheetsHelper) {
      if (String(error?.message || "").includes("Failed to fetch")) {
        sheetsHelper.textContent = "This sheet is likely private. Publish it or share a public CSV export link.";
      } else {
        sheetsHelper.textContent = "Make the sheet public or use a shared CSV export link.";
      }
    }
    if (String(error?.message || "").includes("Failed to fetch")) {
      showError("Request failed. Confirm Sheets API enabled and origin added in Google Cloud.", `Details: ${error.message}`);
      return;
    }
    showError("Google Sheets CSV could not be loaded. Check the link and sharing settings.", `Details: ${error.message}`);
  }
}

function valuesToCsv(values) {
  if (!values.length) return "";
  return values.map((row) => row.map(csvEscape).join(",")).join("\n");
}

async function handlePdfUpload(file) {
  if (!file) return;
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
    const rows = [];

    for (let pageNum = 1; pageNum <= maxPages; pageNum += 1) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const items = (textContent.items || [])
        .filter((item) => String(item.str || "").trim() !== "")
        .map((item) => ({
          text: String(item.str || "").trim(),
          x: item.transform[4],
          y: item.transform[5],
        }));

      const rowMap = new Map();
      items.forEach((item) => {
        const yKey = Math.round(item.y / 3) * 3;
        if (!rowMap.has(yKey)) rowMap.set(yKey, []);
        rowMap.get(yKey).push(item);
      });

      const pageRows = Array.from(rowMap.entries())
        .sort((a, b) => b[0] - a[0])
        .map(([, rowItems]) => {
          const sorted = rowItems.sort((a, b) => a.x - b.x);
          const cells = [];
          let current = sorted[0]?.text || "";
          let lastX = sorted[0]?.x ?? 0;
          for (let i = 1; i < sorted.length; i += 1) {
            const gap = sorted[i].x - lastX;
            if (gap > 20) {
              cells.push(current);
              current = sorted[i].text;
            } else {
              current = `${current} ${sorted[i].text}`.trim();
            }
            lastX = sorted[i].x;
          }
          if (current) cells.push(current);
          return cells;
        })
        .filter((cells) => cells.length >= 2);

      rows.push(...pageRows);
    }

    if (rows.length < 2) {
      showError("No table like content detected. Export a CSV or use a PDF with a clear table.");
      if (pdfStatus) pdfStatus.textContent = "PDF ingestion failed";
      return;
    }

    const csvText = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
    parseCsvText(csvText);
    if (pdfStatus) pdfStatus.textContent = "PDF parsed. Generating dashboard...";
    switchTab("upload");
  } catch (error) {
    showError("No table like content detected. Export a CSV or use a PDF with a clear table.", `Details: ${error.message}`);
    if (pdfStatus) pdfStatus.textContent = "PDF ingestion failed";
  }
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

  if (!state.schema.numeric.length) {
    const detail = buildNumericFailureDetail(state.schema.profiles);
    const message = numericCandidates.length
      ? "No usable numeric metrics detected. Add at least one numeric column."
      : "No usable numeric metrics found. We detected numeric identifiers only. Add metrics like revenue, spend, clicks, orders.";
    showError(message, detail);
    return;
  }

  const recommendedMetrics = chooseKpiMetrics(state.schema.profiles, state.schema.numeric);
  state.selections.primaryMetric = recommendedMetrics[0] || state.schema.numeric[0] || null;
  state.selectedMetric = state.selections.primaryMetric;
  state.selectedDimension = chooseBestDimension(state.schema.profiles, state.schema.categoricals) || state.schema.dates[0] || null;

  initControls(recommendedMetrics);
  setStatus("Rendering dashboard");
  applyFiltersAndRender();

  datasetSummary.textContent = `${rows.length} rows · ${state.schema.columns.length} columns`;
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
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return Number.isNaN(value.valueOf()) ? null : value;
  const raw = String(value).trim();
  if (!raw) return null;

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
    const uniqueRate = nonBlankCount ? uniqueCount / nonBlankCount : 0;

    const stats = {
      type,
      missingRate,
      uniqueCount,
      uniqueRate,
      nonNumericRate: type === "numeric" ? 1 - numericRate : null,
      numericRate,
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
  const numericCandidates = (state.debug.numericCandidates || []).join(", ") || "—";
  const usableNumeric = (state.debug.usableNumericMetrics || []).join(", ") || "—";
  const dateCandidates = (state.debug.dateCandidates || []).join(", ") || "—";
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
  if (rowCount >= 20 && uniqueRate >= 0.9) return true;
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
    option.textContent = metric;
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
  const scopedRows = getFilteredRows();
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
  if (!state.schema.numeric.length) {
    const message = numericCandidates.length
      ? "No usable numeric metrics detected. Add at least one numeric column."
      : "No usable numeric metrics found. We detected numeric identifiers only. Add metrics like revenue, spend, clicks, orders.";
    showError(message);
    return;
  }
  if (!state.selections.primaryMetric || !state.schema.numeric.includes(state.selections.primaryMetric)) {
    state.selections.primaryMetric = state.schema.numeric[0] || null;
  }
  state.selectedMetric = state.selections.primaryMetric;
  state.selections.compareMetrics = (state.selections.compareMetrics || []).filter((m) => state.schema.numeric.includes(m));

  if (filterBadge) {
    filterBadge.textContent = `Filtered to: ${state.filters.industry || "All"}`;
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
}

function runStep(step, fn) {
  try {
    console.log(`STEP: ${step}`);
    fn();
    return true;
  } catch (error) {
    console.error(`STEP FAILED: ${step}`, error);
    showError(`Dashboard step failed: ${step}.`, `Details: ${error?.message || error}`);
    return false;
  }
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
  kpiGrid.innerHTML = "";
  const chosenMetrics = chooseKpiMetrics(state.schema.profiles, state.schema.numeric);
  const metrics = (chosenMetrics && chosenMetrics.length ? chosenMetrics : (state.schema.numeric || []).slice(0, 6)) || [];

  (metrics || []).forEach((metric) => {
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

  const metric = state.selections.primaryMetric || chooseTopNumericByVariance(state.schema.profiles, state.schema.numeric || []);
  state.selections.primaryMetric = metric;
  state.selectedMetric = metric;
  const metricValues = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  const metricType = inferMetricType(metric, metricValues);

  if (state.dateColumn) {
    state.chosenXAxisType = "date";
    const bucket = rows.length > 200 ? "week" : "day";
    const series = aggregateByDate(rows, state.dateColumn, metric, bucket, state.dateRange);
    trendTitle.textContent = `${metric} trend`;
    trendSubtitle.textContent = `Grouped by ${bucket}`;
    trendChartInstance = createLineChart("trendChart", series.labels, series.values, metricType);
  } else {
    const category = chooseCategoryColumn(rows, state.schema.profiles);
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

  let dimension = state.selectedDimension || chooseBestDimension(state.schema.profiles, state.schema.categoricals || []) || state.dateColumn;
  let breakdown = null;
  if (dimension) {
    state.selectedDimension = dimension;
    if (dimensionSelect.value !== dimension) dimensionSelect.value = dimension;
    breakdown = aggregateByCategory(rows, dimension, metric, state.topN);
    barTitle.textContent = `${metric} by ${dimension}`;
    barSubtitle.textContent = `Top ${state.topN}`;
    barChartInstance = createBarChart("barChart", breakdown.labels, breakdown.values, metricType);
  } else {
    state.selectedDimension = null;
    if (dimensionSelect) dimensionSelect.value = "";
    const labels = rows.map((_, index) => index + 1);
    const values = rows.map((row) => parseNumber(row[metric]) ?? 0);
    barTitle.textContent = `${metric} by row`;
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
    extraChartInstance = createHistogram("extraChart", metricValues);
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
  const lower = metric.toLowerCase();
  const isCurrency = /(revenue|sales|cost|spend|profit|price|amount)/i.test(lower);
  const isRateName = /(rate|percent|ctr|cvr|roas|ratio)/i.test(lower);
  const isRateRange = values.length > 0 && values.every((val) => val >= 0 && val <= 1);
  const isRate = isRateName || isRateRange;
  return { isCurrency, isRate };
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

function computeCoefficientOfVariation(rows, metric) {
  const values = rows.map((row) => parseNumber(row[metric])).filter((value) => value !== null);
  if (values.length < 2) return null;
  const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
  if (mean === 0) return null;
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  const std = Math.sqrt(variance);
  return std / Math.abs(mean);
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
  (rows || []).forEach((row) => {
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
  if (!dimension) {
    return { labels: [], values: [] };
  }
  const totals = new Map();
  (rows || []).forEach((row) => {
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
            label: (context) => formatMetricValue(context.parsed.y, metricType),
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
        y: { beginAtZero: true },
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
  } else {
    warnings.push("No date column detected. Charts use categories or row order.");
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
  const insights = buildDecisionInsights(rows);
  insights.forEach((insight) => {
    const card = document.createElement("div");
    card.className = "comparison-card";
    card.innerHTML = `
      <strong>${insight.title}</strong>
      <p class="helper-text">${insight.summary}</p>
      <ul>
        ${insight.bullets.map((bullet) => `<li>${bullet}</li>`).join("")}
      </ul>
    `;
    insightsList.appendChild(card);
  });
}

function buildDecisionInsights(rows) {
  const insights = [];
  const metric = state.selections.primaryMetric;
  if (!metric) return insights;
  const contextPrefix = state.filters.industry !== "All" ? `Within ${state.filters.industry}, ` : "";

  const segmentColumn = state.industryColumn || chooseCategoryColumn(rows, state.schema.profiles);
  if (segmentColumn) {
    const breakdown = aggregateByCategory(rows, segmentColumn, metric, 12);
    if (breakdown.labels.length >= 2) {
      const top = { label: breakdown.labels[0], value: breakdown.values[0] };
      const bottom = { label: breakdown.labels[breakdown.labels.length - 1], value: breakdown.values[breakdown.values.length - 1] };
      const delta = top.value - bottom.value;
      const pct = bottom.value ? delta / bottom.value : 0;
      insights.push({
        title: `${contextPrefix}High vs low performing segment`,
        summary: `${top.label} leads while ${bottom.label} trails on ${metric}.`,
        bullets: [
          `Top: ${top.label} (${formatMetricValue(top.value, inferMetricType(metric, [top.value]))})`,
          `Bottom: ${bottom.label} (${formatMetricValue(bottom.value, inferMetricType(metric, [bottom.value]))}), Δ ${formatMetricValue(delta, inferMetricType(metric, [delta]))} (${percentFormatter.format(pct)})`,
        ],
      });
    }
  }

  if (state.dateColumn) {
    const series = aggregateByDate(rows, state.dateColumn, metric, "day", state.dateRange);
    const window = series.values.length >= 14 ? 7 : series.values.length >= 6 ? 3 : 0;
    if (window) {
      const recent = series.values.slice(-window);
      const prior = series.values.slice(-window * 2, -window);
      const recentTotal = recent.reduce((sum, val) => sum + val, 0);
      const priorTotal = prior.reduce((sum, val) => sum + val, 0);
      const delta = recentTotal - priorTotal;
      const pct = priorTotal ? delta / priorTotal : 0;
      insights.push({
        title: `${contextPrefix}Period over period change`,
        summary: `${metric} is ${delta >= 0 ? "up" : "down"} versus the prior period.`,
        bullets: [
          `Last ${window}: ${formatMetricValue(recentTotal, inferMetricType(metric, [recentTotal]))}`,
          `Change: ${formatMetricValue(delta, inferMetricType(metric, [delta]))} (${percentFormatter.format(pct)})`,
        ],
      });
    }
  } else {
    const category = chooseCategoryColumn(rows, state.schema.profiles);
    if (category) {
      const breakdown = aggregateByCategory(rows, category, metric, 5);
      insights.push({
        title: `${contextPrefix}Top categories contributing most`,
        summary: `${breakdown.labels[0]} leads ${metric} among ${category}.`,
        bullets: [
          `Top: ${breakdown.labels[0]} (${formatMetricValue(breakdown.values[0], inferMetricType(metric, [breakdown.values[0]]))})`,
          `Second: ${breakdown.labels[1] || "—"} (${breakdown.values[1] ? formatMetricValue(breakdown.values[1], inferMetricType(metric, [breakdown.values[1]])) : "—"})`,
        ],
      });
    }
  }

  if (segmentColumn) {
    const breakdown = aggregateByCategory(rows, segmentColumn, metric, 8);
    if (breakdown.labels.length) {
      const total = breakdown.values.reduce((sum, val) => sum + val, 0);
      const share = total ? breakdown.values[0] / total : 0;
      insights.push({
        title: `${contextPrefix}Concentration`,
        summary: `${breakdown.labels[0]} contributes ${percentFormatter.format(share)} of ${metric}.`,
        bullets: [
          `Top segment share: ${percentFormatter.format(share)}`,
          share > 0.4 ? "High concentration (>40%)" : "Balanced distribution",
        ],
      });
    }
  }

  const variance = computeCoefficientOfVariation(rows, metric);
  if (variance !== null) {
    insights.push({
      title: `${contextPrefix}Volatility`,
      summary: variance > 0.35 ? "Metric is volatile across segments/time." : "Metric is stable relative to its mean.",
      bullets: [
        `Coefficient of variation: ${(variance * 100).toFixed(1)}%`,
        variance > 0.35 ? "Consider investigating drivers of swings." : "Performance is consistent.",
      ],
    });
  }

  if (state.schema.numeric.length >= 2) {
    const candidates = state.schema.numeric.filter((m) => m !== metric);
    let best = null;
    candidates.forEach((candidate) => {
      const corr = computeCorrelation(rows, metric, candidate);
      if (corr === null) return;
      if (!best || Math.abs(corr) > Math.abs(best.corr)) {
        best = { metric: candidate, corr };
      }
    });
    if (best) {
      insights.push({
        title: `${contextPrefix}Driver candidate`,
        summary: `Strongest relationship is with ${best.metric}.`,
        bullets: [
          `Correlation: ${best.corr.toFixed(2)}`,
          `Direction: ${best.corr >= 0 ? "Positive" : "Negative"}`,
        ],
      });
    }
  }

  const riskColumn = Object.entries(state.schema.profiles || {})
    .map(([col, stats]) => ({ col, missingRate: stats.missingRate || 0 }))
    .filter((item) => item.missingRate >= 0.3)
    .sort((a, b) => b.missingRate - a.missingRate)[0];
  if (riskColumn) {
    insights.push({
      title: `${contextPrefix}Data quality risk`,
      summary: `${riskColumn.col} has high missing data that could distort insights.`,
      bullets: [
        `Missing rate: ${(riskColumn.missingRate * 100).toFixed(0)}%`,
        "Consider validating data collection or filtering missing rows.",
      ],
    });
  }

  return insights.slice(0, 6);
}

function renderMetricComparison(rows) {
  if (!rows.length) return;
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
    const selected = (state.selections.compareMetrics || []).length
      ? state.selections.compareMetrics
      : chooseTopNumericByVariance(state.schema.profiles, availableCompare, 3);
    state.selections.compareMetrics = Array.isArray(selected) ? selected : [];
    compareMetricSelect.innerHTML = "";
    (availableCompare || []).forEach((metric) => {
      const option = document.createElement("option");
      option.value = metric;
      option.textContent = metric;
      option.selected = selected.includes(metric);
      compareMetricSelect.appendChild(option);
    });
    compareMetricSelect.onchange = () => {
      const chosen = Array.from(compareMetricSelect.selectedOptions).map((opt) => opt.value).slice(0, 5);
      state.selections.compareMetrics = Array.isArray(chosen) ? chosen : [];
      applyFiltersAndRender();
    };
  }

  renderComparisonTrend(rows, primary);
  renderDriversList(rows, primary, state.selections.compareMetrics);
  renderSegmentView(rows, primary);
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

function renderDriversList(rows, primary, compareMetrics) {
  if (!driversList) return;
  driversList.innerHTML = "";
  const metrics = Array.isArray(compareMetrics) ? compareMetrics : [];
  if (!metrics.length) {
    driversList.innerHTML = "<li>Select metrics to compare.</li>";
    return;
  }

  const items = metrics.map((metric) => {
    const corr = computeCorrelation(rows, primary, metric);
    if (corr === null) {
      return { metric, corr: null, label: "Insufficient data" };
    }
    const strength = Math.abs(corr) >= 0.6 ? "Strong" : Math.abs(corr) >= 0.3 ? "Moderate" : "Weak";
    const direction = corr >= 0 ? "positive" : "negative";
    return { metric, corr, label: `${strength} ${direction}` };
  }).sort((a, b) => (b.corr ?? 0) - (a.corr ?? 0));

  items.forEach((item) => {
    const li = document.createElement("li");
    if (item.corr === null) {
      li.textContent = `${item.metric}: insufficient data`;
    } else {
      li.textContent = `${item.metric}: ${item.label} (r = ${item.corr.toFixed(2)})`;
    }
    driversList.appendChild(li);
  });
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
