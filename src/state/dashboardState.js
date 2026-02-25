export function createDashboardState(initial = {}) {
  return {
    dataset: initial.dataset || null,
    selectedMetric: initial.selectedMetric || null,
    selectedDimension: initial.selectedDimension || null,
    chartType: initial.chartType || "bar",
    filters: Array.isArray(initial.filters) ? initial.filters : [],
    sort: initial.sort || null,
    topN: Number.isFinite(initial.topN) ? initial.topN : 8,
    aggregation: initial.aggregation || "sum",
    timeField: initial.timeField || null,
    history: {
      past: [],
      future: [],
    },
  };
}

export function stripHistory(state) {
  const { history, ...rest } = state || {};
  return { ...rest };
}

export function withHistory(state, nextState) {
  const past = [...(state?.history?.past || []), stripHistory(state)];
  return {
    ...nextState,
    history: {
      past,
      future: [],
    },
  };
}
