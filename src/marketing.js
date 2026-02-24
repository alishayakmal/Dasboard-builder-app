import { renderSampleDashboard } from "./marketingSampleDashboard.js";

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => renderSampleDashboard());
} else {
  renderSampleDashboard();
}
