import { renderSampleInsightsSection } from "./components/marketing/sampleInsightsSection.js";
import { renderMiniDashboardPreview } from "./components/marketing/miniDashboardPreview.js";

const insightsRoot = document.getElementById("sampleInsightsRoot");
const miniDashboardRoot = document.getElementById("miniDashboardRoot");

if (insightsRoot) {
  renderSampleInsightsSection(insightsRoot);
}

if (miniDashboardRoot) {
  renderMiniDashboardPreview(miniDashboardRoot);
}
