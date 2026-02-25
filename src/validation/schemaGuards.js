const numericTypes = new Set(["number", "currency", "percent", "numeric"]);
const categoricalTypes = new Set(["string", "categorical"]);
const dateTypes = new Set(["date", "datetime", "time"]);

export function buildFieldMap(datasetSchema = {}) {
  const map = new Map();
  const columns = Array.isArray(datasetSchema.columns) ? datasetSchema.columns : [];
  columns.forEach((col) => {
    if (!col?.name) return;
    const type = String(col.type || "").toLowerCase() || "string";
    map.set(col.name, type);
  });
  return map;
}

function fieldType(fieldMap, field) {
  return fieldMap.get(field) || "string";
}

export function validateMetric(schema, state, field) {
  const fieldMap = buildFieldMap(schema);
  if (!fieldMap.has(field)) {
    return { ok: false, code: "metric_not_found", message: "That metric is not in the dataset." };
  }
  const type = fieldType(fieldMap, field);
  const agg = state?.aggregation || "sum";
  if ((agg === "sum" || agg === "avg") && !numericTypes.has(type)) {
    return { ok: false, code: "metric_not_numeric", message: "That metric is not numeric so avg is unavailable." };
  }
  return { ok: true };
}

export function validateAggregation(schema, state, aggregation) {
  const fieldMap = buildFieldMap(schema);
  const metric = state?.selectedMetric;
  if (metric && !fieldMap.has(metric)) {
    return { ok: false, code: "metric_not_found", message: "That metric is not in the dataset." };
  }
  const type = metric ? fieldType(fieldMap, metric) : "string";
  if ((aggregation === "sum" || aggregation === "avg") && !numericTypes.has(type)) {
    return { ok: false, code: "agg_requires_numeric", message: "That metric is not numeric so avg is unavailable." };
  }
  if (aggregation === "count_distinct" && !metric) {
    return { ok: false, code: "metric_required", message: "Select a metric before using count distinct." };
  }
  return { ok: true };
}

export function validateChartType(schema, state, chartType) {
  const timeField = state?.timeField || schema?.dateField || null;
  if ((chartType === "line" || chartType === "area") && !timeField) {
    return { ok: false, code: "missing_time_field", message: "Time series charts require a date field." };
  }
  return { ok: true };
}

export function validateFilter(schema, filter) {
  const fieldMap = buildFieldMap(schema);
  if (!fieldMap.has(filter.field)) {
    return { ok: false, code: "filter_field_missing", message: "That filter field is not in the dataset." };
  }
  const type = fieldType(fieldMap, filter.field);
  const operator = String(filter.operator || "").toLowerCase();
  if (numericTypes.has(type)) {
    const allowed = new Set(["=", "==", "equals", "!=", ">", "<", ">=", "<=", "between"]);
    if (!allowed.has(operator)) {
      return { ok: false, code: "invalid_filter_operator", message: "Numeric filters must use comparison operators." };
    }
    return { ok: true };
  }
  if (dateTypes.has(type)) {
    const allowed = new Set(["=", "==", "equals", "on", "before", "after", "between"]);
    if (!allowed.has(operator)) {
      return { ok: false, code: "invalid_filter_operator", message: "Date filters must use before/after/on/between." };
    }
    return { ok: true };
  }
  const allowed = new Set(["=", "==", "equals", "!=", "contains", "in"]);
  if (!allowed.has(operator)) {
    return { ok: false, code: "invalid_filter_operator", message: "Text filters must use equals or contains." };
  }
  return { ok: true };
}

export function validateToolCall(schema, state, toolCall) {
  const name = toolCall?.name;
  const args = toolCall?.args || {};
  switch (name) {
    case "setMetric":
      return validateMetric(schema, state, args.field);
    case "setDimension":
      return buildFieldMap(schema).has(args.field)
        ? { ok: true }
        : { ok: false, code: "dimension_not_found", message: "That dimension is not in the dataset." };
    case "setChartType":
      return validateChartType(schema, state, args.type);
    case "setAggregation":
      return validateAggregation(schema, state, args.agg);
    case "applyFilter":
      return validateFilter(schema, { field: args.field, operator: args.operator, value: args.value });
    case "clearFilter":
      return { ok: true };
    case "setTopN":
      return Number.isFinite(args.n) && args.n > 0
        ? { ok: true }
        : { ok: false, code: "invalid_topn", message: "Top N must be a positive number." };
    case "setSort":
      return args.field ? { ok: true } : { ok: false, code: "invalid_sort", message: "Sort requires a field." };
    case "undo":
    case "redo":
      return { ok: true };
    default:
      return { ok: false, code: "unknown_action", message: `Unknown action: ${name}` };
  }
}
