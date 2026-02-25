const OpenAI = require("openai");

const ACTIONS = new Set([
  "setMetric",
  "setDimension",
  "setChartType",
  "setAggregation",
  "applyFilter",
  "clearFilter",
  "setTopN",
  "setSort",
]);

function buildSystemPrompt() {
  return [
    "You are a chart editing assistant.",
    "You MUST return only valid JSON with this schema:",
    '{"toolCalls":[{"name":"setMetric","args":{"field":"<column>"}}],"assistantMessage":"...","proposedDiff":{}}',
    "Do not include markdown, commentary, or extra keys.",
    "Only use action names from the allow list.",
    "Only use column names that exist in datasetSchema.columns[].name.",
  ].join(" ");
}

function buildUserPrompt({ message, datasetSchema, currentDashboardState }) {
  return [
    `User request: ${message}`,
    `Dataset schema: ${JSON.stringify(datasetSchema || {})}`,
    `Current dashboard state: ${JSON.stringify(currentDashboardState || {})}`,
    "Return JSON only.",
  ].join("\n");
}

function validateToolCalls(toolCalls) {
  if (!Array.isArray(toolCalls)) return { ok: false, error: "toolCalls must be an array." };
  for (const call of toolCalls) {
    if (!ACTIONS.has(call?.name)) {
      return { ok: false, error: `Unknown action: ${call?.name}` };
    }
  }
  return { ok: true };
}

async function generateChatPlan({ message, datasetSchema, currentDashboardState }) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return { ok: false, error: "Missing OPENAI_API_KEY" };
  }

  const client = new OpenAI({ apiKey });
  const model = process.env.OPENAI_MODEL || "gpt-4o-mini";
  const response = await client.chat.completions.create({
    model,
    temperature: 0.2,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: buildSystemPrompt() },
      { role: "user", content: buildUserPrompt({ message, datasetSchema, currentDashboardState }) },
    ],
  });

  const content = response?.choices?.[0]?.message?.content || "";
  let parsed = null;
  try {
    parsed = JSON.parse(content);
  } catch (error) {
    return { ok: false, error: "AI returned invalid JSON.", detail: content.slice(0, 200) };
  }

  const validation = validateToolCalls(parsed.toolCalls);
  if (!validation.ok) {
    return { ok: false, error: validation.error };
  }

  return {
    ok: true,
    toolCalls: parsed.toolCalls,
    assistantMessage: parsed.assistantMessage || "Proposed updates ready for review.",
    proposedDiff: parsed.proposedDiff || {},
    notes: parsed.notes || "",
    model,
  };
}

module.exports = { generateChatPlan };
