import { sampleInsights } from "../data/sampleInsights.js";
import { miniKpis, miniTrend } from "../data/sampleMiniDashboard.js";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export async function getSampleInsights() {
  await delay(400);
  return [...sampleInsights];
}

export async function getMiniDashboard() {
  await delay(400);
  return { miniKpis: [...miniKpis], miniTrend: [...miniTrend] };
}
