/**
 * Canonical dashboard renderer contract (plain JS app version).
 * All source ingest paths should first produce a NormalizedDataset and then call this renderer.
 *
 * @typedef {{
 * rows: Record<string, any>[],
 * columns: { name: string, type: "number" | "currency" | "percent" | "date" | "string" }[],
 * meta: { sourceType: "csv" | "pdf" | "sheet" | "api" | "demo", name?: string, dateField?: string, pdfMode?: "table" | "text" }
 * }} NormalizedDataset
 */

/**
 * @param {{ dataset: NormalizedDataset, mode: "demo" | "user" | "pdf-text", delegate: Function }} args
 */
export function DashboardRenderer({ dataset, mode, delegate }) {
  if (!dataset || typeof delegate !== "function") return;
  delegate({ dataset, mode });
}
