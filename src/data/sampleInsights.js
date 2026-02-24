/**
 * @typedef {"Growth" | "Risk" | "Opportunity" | "Efficiency"} InsightCategory
 * @typedef {"percent" | "percentagePoints"} DeltaType
 * @typedef {"line" | "bar" | "table"} ChartType
 * @typedef {"low" | "medium" | "high"} SupportLevel
 *
 * @typedef {Object} InsightMetric
 * @property {string} name
 * @property {string} unit
 * @property {string} definition
 * @property {(value:number) => string} format
 *
 * @typedef {Object} InsightPeriod
 * @property {string} currentPeriodStart
 * @property {string} currentPeriodEnd
 * @property {string} comparisonPeriodStart
 * @property {string} comparisonPeriodEnd
 * @property {"WoW" | "MoM" | "QoQ" | "YoY"} comparisonType
 *
 * @typedef {Object} InsightEvidence
 * @property {ChartType} primaryChartType
 * @property {string} breakdownDimension
 * @property {{ total:number, byGroup: Record<string, number> }} sampleSize
 * @property {Array<{ x:string, y:number, group?:string }>} chartPoints
 *
 * @typedef {Object} InsightConfidence
 * @property {SupportLevel} supportLevel
 * @property {string} reason
 *
 * @typedef {Object} InsightAction
 * @property {string} text
 *
 * @typedef {Object} InsightDrilldown
 * @property {string[]} topSegments
 * @property {Array<Record<string, string|number>>} rows
 *
 * @typedef {Object} SampleInsight
 * @property {string} id
 * @property {string} title
 * @property {InsightCategory} category
 * @property {InsightMetric} metric
 * @property {InsightPeriod} period
 * @property {{ currentValue:number, deltaAbsolute:number, deltaRelative:number, deltaType:DeltaType }} value
 * @property {InsightEvidence} evidence
 * @property {InsightConfidence} confidence
 * @property {InsightAction[]} actions
 * @property {InsightDrilldown} drilldown
 */

const percent = (value) => `${(value * 100).toFixed(1)}%`;
const currency = (value) => `$${value.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ",")}`;

/** @type {SampleInsight[]} */
export const sampleInsights = [
  {
    id: "revenue-acceleration",
    title: "Revenue acceleration (Enterprise vs Mid‑market)",
    category: "Growth",
    metric: {
      name: "Revenue",
      unit: "USD",
      definition: "Sum of Revenue over the current period.",
      format: currency,
    },
    period: {
      currentPeriodStart: "2024-02-01",
      currentPeriodEnd: "2024-02-07",
      comparisonPeriodStart: "2024-01-25",
      comparisonPeriodEnd: "2024-01-31",
      comparisonType: "WoW",
    },
    value: {
      currentValue: 1285000,
      deltaAbsolute: 145000,
      deltaRelative: 0.13,
      deltaType: "percent",
    },
    evidence: {
      primaryChartType: "line",
      breakdownDimension: "plan",
      sampleSize: {
        total: 124,
        byGroup: { Enterprise: 62, "Mid-market": 44, SMB: 18 },
      },
      chartPoints: [
        { x: "2024-01-11", y: 320000, group: "Enterprise" },
        { x: "2024-01-18", y: 342000, group: "Enterprise" },
        { x: "2024-01-25", y: 360000, group: "Enterprise" },
        { x: "2024-02-01", y: 418000, group: "Enterprise" },
        { x: "2024-01-11", y: 210000, group: "Mid-market" },
        { x: "2024-01-18", y: 198000, group: "Mid-market" },
        { x: "2024-01-25", y: 205000, group: "Mid-market" },
        { x: "2024-02-01", y: 214000, group: "Mid-market" },
      ],
    },
    confidence: {
      supportLevel: "high",
      reason: "4 periods with consistent upward slope in Enterprise revenue.",
    },
    actions: [
      { text: "Review Enterprise pipeline stages to confirm acceleration is driven by late‑stage deals." },
      { text: "Compare mid‑market discounting to identify margin impact." },
      { text: "Validate plan lift = Enterprise growth rate minus Mid‑market growth rate." },
    ],
    drilldown: {
      topSegments: ["Enterprise", "Mid-market"],
      rows: [
        { plan: "Enterprise", revenue: 418000, week: "2024-02-01" },
        { plan: "Mid-market", revenue: 214000, week: "2024-02-01" },
        { plan: "Enterprise", revenue: 360000, week: "2024-01-25" },
      ],
    },
  },
  {
    id: "channel-mix-shift",
    title: "Channel mix shift (Pipeline share)",
    category: "Opportunity",
    metric: {
      name: "Pipeline share",
      unit: "percent",
      definition: "Pipeline amount by channel / total pipeline amount.",
      format: percent,
    },
    period: {
      currentPeriodStart: "2024-02-01",
      currentPeriodEnd: "2024-02-07",
      comparisonPeriodStart: "2024-01-25",
      comparisonPeriodEnd: "2024-01-31",
      comparisonType: "WoW",
    },
    value: {
      currentValue: 0.42,
      deltaAbsolute: 0.06,
      deltaRelative: 0.17,
      deltaType: "percentagePoints",
    },
    evidence: {
      primaryChartType: "line",
      breakdownDimension: "channel",
      sampleSize: {
        total: 98,
        byGroup: { "Paid Search": 38, "Paid Social": 24, "Organic": 36 },
      },
      chartPoints: [
        { x: "2024-01-11", y: 0.34, group: "Paid Search" },
        { x: "2024-01-18", y: 0.36, group: "Paid Search" },
        { x: "2024-01-25", y: 0.36, group: "Paid Search" },
        { x: "2024-02-01", y: 0.42, group: "Paid Search" },
      ],
    },
    confidence: {
      supportLevel: "medium",
      reason: "Change exceeds 5 pp with >30 rows total.",
    },
    actions: [
      { text: "Inspect Paid Search keyword mix to confirm pipeline quality." },
      { text: "Rebalance spend if Paid Search conversion rate remains higher than Paid Social." },
      { text: "Monitor Organic share to avoid over‑reliance on one channel." },
    ],
    drilldown: {
      topSegments: ["Paid Search", "Organic"],
      rows: [
        { channel: "Paid Search", pipeline: 420000, week: "2024-02-01" },
        { channel: "Organic", pipeline: 280000, week: "2024-02-01" },
        { channel: "Paid Social", pipeline: 210000, week: "2024-02-01" },
      ],
    },
  },
];
