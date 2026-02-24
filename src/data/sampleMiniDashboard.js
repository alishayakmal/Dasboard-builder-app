/**
 * @typedef {Object} MiniPoint
 * @property {number} x
 * @property {number} y
 */

export const miniKpis = [
  {
    label: "Revenue",
    value: "$1.4M",
    delta: "+12%",
    deltaType: "percent",
    window: "Last 7 days vs prior 7 days",
    definition: "Sum of Revenue over the selected period.",
  },
  {
    label: "Retention",
    value: "52%",
    delta: "+3 pp",
    deltaType: "percentagePoints",
    window: "Last 7 days vs prior 7 days",
    definition: "Retained users / eligible users.",
  },
  {
    label: "Active users",
    value: "84k",
    delta: "+9%",
    deltaType: "percent",
    window: "Last 7 days vs prior 7 days",
    definition: "Unique users with at least one session.",
  },
];

/** @type {MiniPoint[]} */
export const miniTrend = [
  { x: 1, y: 42 },
  { x: 2, y: 46 },
  { x: 3, y: 44 },
  { x: 4, y: 50 },
  { x: 5, y: 55 },
  { x: 6, y: 53 },
  { x: 7, y: 58 },
  { x: 8, y: 61 },
  { x: 9, y: 60 },
  { x: 10, y: 66 },
  { x: 11, y: 69 },
  { x: 12, y: 72 },
];
