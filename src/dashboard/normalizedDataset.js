/**
 * @typedef {{
 * rows: Record<string, any>[],
 * columns: { name: string, type: "number" | "currency" | "percent" | "date" | "string" }[],
 * meta: { sourceType: "csv" | "pdf" | "sheet" | "api" | "demo", name?: string, dateField?: string, hasDateField: boolean, pdfMode?: "table" | "text" }
 * }} NormalizedDataset
 */

/**
 * @param {{ rows: Record<string, any>[], columns: { name: string, type: string }[], meta: Record<string, any> }} input
 * @returns {NormalizedDataset}
 */
export function createNormalizedDataset(input) {
  return {
    rows: Array.isArray(input?.rows) ? input.rows : [],
    columns: Array.isArray(input?.columns) ? input.columns : [],
    meta: {
      sourceType: input?.meta?.sourceType || "csv",
      name: input?.meta?.name,
      dateField: input?.meta?.dateField,
      hasDateField: Boolean(
        input?.meta?.hasDateField
        || (Array.isArray(input?.columns) && input.columns.some((column) => column?.type === "date"))
      ),
      pdfMode: input?.meta?.pdfMode,
    },
  };
}
