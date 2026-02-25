import { stripHistory, withHistory } from "../state/dashboardState.js";

function update(state, patch) {
  return withHistory(state, { ...state, ...patch });
}

export function setMetric(state, field) {
  return update(state, { selectedMetric: field });
}

export function setDimension(state, field) {
  return update(state, { selectedDimension: field });
}

export function setChartType(state, type) {
  return update(state, { chartType: type });
}

export function setAggregation(state, aggregation) {
  return update(state, { aggregation });
}

export function applyFilter(state, field, operator, value) {
  const filters = Array.isArray(state.filters) ? [...state.filters] : [];
  const next = { field, operator, value };
  const idx = filters.findIndex((item) => item.field === field);
  if (idx >= 0) filters[idx] = next;
  else filters.push(next);
  return update(state, { filters });
}

export function clearFilter(state, field) {
  const filters = Array.isArray(state.filters) ? state.filters.filter((item) => item.field !== field) : [];
  return update(state, { filters });
}

export function setTopN(state, n) {
  return update(state, { topN: n });
}

export function setSort(state, field, direction) {
  return update(state, { sort: field ? { field, direction } : null });
}

export function undo(state) {
  const past = state?.history?.past || [];
  if (!past.length) return state;
  const previous = past[past.length - 1];
  const future = [stripHistory(state), ...(state?.history?.future || [])];
  return {
    ...previous,
    history: {
      past: past.slice(0, -1),
      future,
    },
  };
}

export function redo(state) {
  const future = state?.history?.future || [];
  if (!future.length) return state;
  const next = future[0];
  const past = [...(state?.history?.past || []), stripHistory(state)];
  return {
    ...next,
    history: {
      past,
      future: future.slice(1),
    },
  };
}
