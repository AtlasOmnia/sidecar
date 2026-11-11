/**
 * Deterministic local arithmetic for selected-range operations.
 *
 * When the user's request is an unambiguous simple arithmetic operation on
 * the current selection (sum, average, min, max, count), this module
 * intercepts the request and computes the result locally using Excel's
 * range statistics rather than making an LLM round trip.
 *
 * No silent writes — this is read-only. The result is returned to the
 * task pane UI as if the model had answered.
 */

import { summarizeRangeValues, RangeStats, parseExcelSelection } from "./excelTools";

/** The set of arithmetic operations we intercept deterministically. */
export type ArithmeticOp = "sum" | "average" | "min" | "max" | "count";

interface DetectedRequest {
  op: ArithmeticOp;
}

/** Patterns that match simple arithmetic requests on selected/current range. */
const PATTERNS: Array<{ regex: RegExp; op: ArithmeticOp }> = [
  // Sum patterns
  { regex: /\b(sum|total|add\s*(up)?)\b/i, op: "sum" },
  { regex: /^what['\u2019]?s?\s+the\s+(sum|total)\b/i, op: "sum" },
  // Average patterns
  { regex: /\b(average|avg|mean)\b/i, op: "average" },
  { regex: /^what['\u2019]?s?\s+the\s+(average|avg|mean)\b/i, op: "average" },
  // Min patterns
  { regex: /\b(minimum|min|lowest|smallest)\b/i, op: "min" },
  { regex: /^what['\u2019]?s?\s+the\s+(minimum|min|lowest|smallest)\b/i, op: "min" },
  // Max patterns
  { regex: /\b(maximum|max|highest|largest|biggest)\b/i, op: "max" },
  { regex: /^what['\u2019]?s?\s+the\s+(maximum|max|highest|largest)\b/i, op: "max" },
  // Count patterns
  { regex: /\b(count|how\s+many)\b/i, op: "count" },
];

/** Non-arithmetic keywords that indicate this is NOT a simple arithmetic request. */
const NON_ARITHMETIC_KEYWORDS = [
  /\b(chart|graph|pivot|table|format|color|bold|font|border|filter|sort|freeze|protect|comment|csv|import|rename|copy|delete|create|add\s+sheet|clear)\b/i,
  /^why|^how\s+(do|can|would|should)/i,
  /\?$/m,  // ends with question mark — probably more complex
];

/**
 * Try to detect a simple arithmetic request on the selection.
 * Returns the detected operation or null if the request is too complex.
 */
export function detectArithmeticRequest(userRequest: string): DetectedRequest | null {
  const trimmed = userRequest.trim().toLowerCase();

  // Skip if too short or too long
  if (trimmed.length < 3 || trimmed.length > 120) return null;

  // Skip if contains any non-arithmetic keywords
  for (const kw of NON_ARITHMETIC_KEYWORDS) {
    if (kw.test(trimmed)) return null;
  }

  // Check for selection/range context words
  const hasSelectionContext =
    /\b(select(ed|ion)?|current|these|this\s+range|highlighted|active)\b/i.test(trimmed) ||
    // Very short requests without explicit range reference are assumed selection
    trimmed.length < 40;

  if (!hasSelectionContext) return null;

  // Find the strongest arithmetic match
  for (const { regex, op } of PATTERNS) {
    if (regex.test(trimmed)) {
      return { op };
    }
  }

  return null;
}

/**
 * Format range statistics as a natural-language response.
 */
export function formatArithmeticResult(
  op: ArithmeticOp,
  stats: RangeStats,
  selectionAddress: string,
): string {
  const { sheet } = parseExcelSelection(selectionAddress, "");
  const sheetInfo = sheet ? ` in ${sheet}` : "";

  switch (op) {
    case "sum":
      return `The sum of the selected range${sheetInfo} (${stats.count} cells, ${stats.numericCount} numeric) is **${formatNumber(stats.sum)}**.`;
    case "average":
      if (stats.average === null) {
        return `The selected range${sheetInfo} has ${stats.count} cells but no numeric values to average.`;
      }
      return `The average of the selected range${sheetInfo} (${stats.numericCount} numeric values) is **${formatNumber(stats.average)}**.`;
    case "min":
      if (stats.min === null) {
        return `The selected range${sheetInfo} has ${stats.count} cells but no numeric values to find a minimum.`;
      }
      return `The minimum value in the selected range${sheetInfo} (${stats.numericCount} numeric values) is **${formatNumber(stats.min)}**.`;
    case "max":
      if (stats.max === null) {
        return `The selected range${sheetInfo} has ${stats.count} cells but no numeric values to find a maximum.`;
      }
      return `The maximum value in the selected range${sheetInfo} (${stats.numericCount} numeric values) is **${formatNumber(stats.max)}**.`;
    case "count":
      return `The selected range${sheetInfo} has **${stats.count}** total cells (${stats.numericCount} numeric, ${stats.blankCount} blank).`;
    default:
      return `Selected range${sheetInfo}: ${stats.count} cells, ${stats.numericCount} numeric, sum=${formatNumber(stats.sum)}, avg=${stats.average !== null ? formatNumber(stats.average) : "N/A"}.`;
  }
}

function formatNumber(value: number): string {
  // Format with reasonable precision — avoid floating-point noise
  if (Number.isInteger(value)) {
    return value.toLocaleString();
  }
  // Up to 4 decimal places, strip trailing zeros
  const rounded = Math.round(value * 10000) / 10000;
  return rounded.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: 4,
  });
}
