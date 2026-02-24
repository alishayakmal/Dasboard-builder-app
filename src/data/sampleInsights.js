/**
 * @typedef {"growth" | "risk" | "opportunity"} InsightSeverity
 * @typedef {Object} SampleInsight
 * @property {string} id
 * @property {string} title
 * @property {string} subtitle
 * @property {string} metricLabel
 * @property {string} metricValue
 * @property {string} metricDelta
 * @property {InsightSeverity} severity
 */

/** @type {SampleInsight[]} */
export const sampleInsights = [
  {
    id: "insight-growth",
    title: "Revenue acceleration",
    subtitle: "Enterprise plan revenue is climbing faster than mid-market.",
    metricLabel: "Plan lift",
    metricValue: "+18%",
    metricDelta: "WoW",
    severity: "growth",
  },
  {
    id: "insight-opportunity",
    title: "Channel mix shift",
    subtitle: "Paid search now drives the largest share of pipeline.",
    metricLabel: "Pipeline share",
    metricValue: "42%",
    metricDelta: "+6 pp",
    severity: "opportunity",
  },
  {
    id: "insight-risk",
    title: "Retention risk",
    subtitle: "Retention dipped below target in EU over the last 2 weeks.",
    metricLabel: "Retention",
    metricValue: "41%",
    metricDelta: "-7 pp",
    severity: "risk",
  },
];
