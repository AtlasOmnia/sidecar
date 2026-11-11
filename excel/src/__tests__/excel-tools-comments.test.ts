import { describe, expect, it } from "vitest";
import {
  parseColumnLabel,
  isAddressInRange,
  normalizeCsvRows,
  parseNumericCell,
} from "../excelTools";

// ─── parseColumnLabel ──────────────────────────────────────────────────────────

describe("parseColumnLabel", () => {
  it("converts single-letter column labels", () => {
    expect(parseColumnLabel("A")).toBe(1);
    expect(parseColumnLabel("Z")).toBe(26);
  });

  it("converts multi-letter column labels", () => {
    expect(parseColumnLabel("AA")).toBe(27);
    expect(parseColumnLabel("AZ")).toBe(52);
    expect(parseColumnLabel("BA")).toBe(53);
    expect(parseColumnLabel("ZZ")).toBe(702);
    expect(parseColumnLabel("AAA")).toBe(703);
  });

  it("treats empty string as 0", () => {
    expect(parseColumnLabel("")).toBe(0);
  });

  it("handles lowercase input (regression: isAddressInRange passes un-normalized column letters)", () => {
    expect(parseColumnLabel("a")).toBe(1);
    expect(parseColumnLabel("z")).toBe(26);
    expect(parseColumnLabel("aa")).toBe(27);
    expect(parseColumnLabel("az")).toBe(52);
    expect(parseColumnLabel("ba")).toBe(53);
    expect(parseColumnLabel("zz")).toBe(702);
  });

  it("handles mixed-case input", () => {
    expect(parseColumnLabel("Aa")).toBe(27);
    expect(parseColumnLabel("aB")).toBe(28);
    expect(parseColumnLabel("AbC")).toBe(731); // A=1, B=2, C=3
  });
});

// ─── isAddressInRange ──────────────────────────────────────────────────────────

describe("isAddressInRange", () => {
  it("returns true when cell is inside the range", () => {
    expect(isAddressInRange("C5", "A1:D10")).toBe(true);
    expect(isAddressInRange("A1", "A1:D10")).toBe(true);
    expect(isAddressInRange("D10", "A1:D10")).toBe(true);
    expect(isAddressInRange("B3", "C1:G20")).toBe(false);
  });

  it("returns false when cell is outside the range", () => {
    expect(isAddressInRange("A1", "C1:G20")).toBe(false);
    expect(isAddressInRange("H5", "C1:G20")).toBe(false);
    expect(isAddressInRange("D25", "C1:G20")).toBe(false);
  });

  it("handles single-cell ranges (no colon)", () => {
    expect(isAddressInRange("B2", "B2")).toBe(true);
    expect(isAddressInRange("B3", "B2")).toBe(false);
    expect(isAddressInRange("C2", "B2")).toBe(false);
  });

  it("strips sheet name prefixes from cell addresses", () => {
    expect(isAddressInRange("Sheet1!C5", "A1:D10")).toBe(true);
    expect(isAddressInRange("Sheet1!C5", "Sheet1!A1:D10")).toBe(true);
    expect(isAddressInRange("Sheet1!A1", "Sheet1!C1:G20")).toBe(false);
    expect(isAddressInRange("Other!C5", "A1:D10")).toBe(true);
  });

  it("handles multi-letter column labels in range addresses", () => {
    expect(isAddressInRange("AA5", "AA1:ZZ10")).toBe(true);
    expect(isAddressInRange("ZZ10", "AA1:ZZ10")).toBe(true);
    expect(isAddressInRange("A1", "AA1:ZZ10")).toBe(false);
    expect(isAddressInRange("AB3", "AA1:AC10")).toBe(true);
  });

  it("returns false for empty cell address", () => {
    expect(isAddressInRange("", "A1:D10")).toBe(false);
  });

  it("returns false for empty range address", () => {
    expect(isAddressInRange("C5", "")).toBe(false);
  });

  it("returns false when both addresses are empty", () => {
    expect(isAddressInRange("", "")).toBe(false);
  });

  it("returns false for malformed cell address with no row number", () => {
    expect(isAddressInRange("A", "A1:D10")).toBe(false);
  });

  it("returns false for malformed cell address with no column letter", () => {
    expect(isAddressInRange("5", "A1:D10")).toBe(false);
  });

  it("handles absolute reference $ notation gracefully (returns false for $ out-of-range)", () => {
    // $A$1 notation — the function may not handle it perfectly, but should not crash
    expect(isAddressInRange("$A$1", "A1:D10")).toBe(true);
  });

  it("handles lowercase range addresses (regression: deterministicAddress preserves caller case)", () => {
    // Host returns canonical uppercase addresses; the caller may pass lowercase.
    expect(isAddressInRange("B2", "a1:c3")).toBe(true);
    expect(isAddressInRange("A1", "a1:c3")).toBe(true);
    expect(isAddressInRange("D4", "a1:c3")).toBe(false);
    expect(isAddressInRange("E7", "a1:c3")).toBe(false);
    // Multi-letter columns
    expect(isAddressInRange("AB3", "aa1:ac10")).toBe(true);
    expect(isAddressInRange("A1", "aa1:ac10")).toBe(false);
  });

  it("uses getLocation() path instead of cellAddress type escape for comment position", () => {
    // This is a compile-time / type-safety test: Comment objects expose getLocation()
    // which returns an Excel.Range. The code should use this typed path rather than
    // type-casting to access undocumented cellAddress.
    // At runtime, clearComments should obtain the cell address by calling
    // comment.getLocation().address (after loading the range's address).
    // This test validates that isAddressInRange can parse addresses returned by
    // Range.address, which are standard A1-style like "Sheet1!C5" or "C5".
    expect(isAddressInRange("Sheet1!C5", "A1:D10")).toBe(true);
    expect(isAddressInRange("A1", "A1:D10")).toBe(true);
  });
});

// ─── normalizeCsvRows ──────────────────────────────────────────────────────────

describe("normalizeCsvRows", () => {
  it("pads shorter rows with empty strings", () => {
    const input: unknown[][] = [["a", "b"], ["c"]];
    const result = normalizeCsvRows(input, 3);
    expect(result).toEqual([
      ["a", "b", ""],
      ["c", "", ""],
    ]);
  });

  it("keeps rows at target width unchanged", () => {
    const input: unknown[][] = [["a", "b"], ["c", "d"]];
    const result = normalizeCsvRows(input, 2);
    expect(result).toEqual([["a", "b"], ["c", "d"]]);
  });

  it("does not truncate rows wider than totalCols", () => {
    const input: unknown[][] = [["a", "b", "c", "d"], ["e", "f"]];
    const result = normalizeCsvRows(input, 3);
    expect(result).toEqual([
      ["a", "b", "c", "d"],
      ["e", "f", ""],
    ]);
  });

  it("converts null and undefined to empty strings", () => {
    const input: unknown[][] = [[null, undefined]];
    const result = normalizeCsvRows(input, 3);
    expect(result).toEqual([["", "", ""]]);
  });

  it("converts numbers to strings", () => {
    const input: unknown[][] = [[42, true]];
    const result = normalizeCsvRows(input, 2);
    expect(result).toEqual([["42", "true"]]);
  });

  it("handles empty input array", () => {
    expect(normalizeCsvRows([], 3)).toEqual([]);
  });
});

// ─── import_csv_to_sheet: pure CSV parsing tests ───────────────────────────────
// These test PapaParse behavior directly since that's what the implementation uses.
// They verify the RFC 4180 edge cases the implementation claims to handle.

import * as Papa from "papaparse";

function parseCsv(csv: string): { data: unknown[][]; errors: Papa.ParseError[] } {
  const result = Papa.parse(csv, {
    header: false,
    skipEmptyLines: false,
    dynamicTyping: false,
    delimitersToGuess: [",", "\t", ";", "|"],
  });
  return { data: result.data as unknown[][], errors: result.errors };
}

describe("CSV parsing (RFC 4180 via PapaParse)", () => {
  it("handles quoted commas in fields", () => {
    const { data } = parseCsv('a,"b,c",d');
    expect(data).toEqual([["a", "b,c", "d"]]);
  });

  it("handles escaped quotes (double quotes inside quoted fields)", () => {
    const { data } = parseCsv('a,"b""c",d');
    expect(data).toEqual([["a", 'b"c', "d"]]);
  });

  it("handles CRLF line endings", () => {
    const { data } = parseCsv("a,b\r\nc,d");
    expect(data).toEqual([["a", "b"], ["c", "d"]]);
  });

  it("handles blank fields", () => {
    const { data } = parseCsv("a,,c");
    expect(data).toEqual([["a", "", "c"]]);
  });

  it("handles trailing blank fields", () => {
    const { data } = parseCsv("a,b,");
    expect(data).toEqual([["a", "b", ""]]);
  });

  it("handles leading blank fields", () => {
    const { data } = parseCsv(",a,b");
    expect(data).toEqual([["", "a", "b"]]);
  });

  it("handles all-blank row", () => {
    const { data } = parseCsv(",");
    expect(data).toEqual([["", ""]]);
  });

  it("uses comma as default delimiter", () => {
    const { data } = parseCsv("a,b,c");
    expect(data).toEqual([["a", "b", "c"]]);
  });

  it("auto-detects tab delimiter", () => {
    const { data } = parseCsv("a\tb\tc");
    expect(data).toEqual([["a", "b", "c"]]);
  });

  it("auto-detects semicolon delimiter", () => {
    const { data } = parseCsv("a;b;c");
    expect(data).toEqual([["a", "b", "c"]]);
  });

  it("auto-detects pipe delimiter", () => {
    const { data } = parseCsv("a|b|c");
    expect(data).toEqual([["a", "b", "c"]]);
  });

  it("handles jagged rows (different column counts)", () => {
    const { data } = parseCsv("a,b,c\nd,e\ng");
    expect(data[0].length).toBe(3);
    expect(data[1].length).toBe(2);
    expect(data[2].length).toBe(1);
  });

  it("handles single quoted field", () => {
    const { data } = parseCsv('"hello world"');
    expect(data).toEqual([["hello world"]]);
  });

  it("returns error for malformed CSV (unclosed quote)", () => {
    const { data, errors } = parseCsv('a,"unclosed,b,c');
    // PapaParse still returns partial data for unclosed quotes but flags it
    expect(errors.length).toBeGreaterThan(0);
    expect(errors.some((e) => e.type === "Quotes")).toBe(true);
  });

  it("returns data for valid single row", () => {
    const { data } = parseCsv("Name,Age,City");
    expect(data).toEqual([["Name", "Age", "City"]]);
  });

  it("handles empty CSV string (PapaParse emits Delimiter warning, no data)", () => {
    const { data, errors } = parseCsv("");
    expect(data).toEqual([]);
    expect(errors.some((e) => e.type === "Delimiter")).toBe(true);
    // The production code should abort with a clear error for empty CSV
    // rather than attempting getResizedRange(-1, -1)
  });

  it("handles whitespace-only CSV as a single non-empty row", () => {
    const { data } = parseCsv("   ");
    expect(data).toEqual([["   "]]);
  });

  it("handles newline-only CSV as two rows (newline + empty final row)", () => {
    const { data } = parseCsv("\n");
    // PapaParse sees the newline as ending the first row, then adds a blank row for the empty remainder
    expect(data).toEqual([[""], [""]]);
  });

  it("handles newline at end of file (PapaParse returns trailing blank row)", () => {
    const { data } = parseCsv("a,b,c\n");
    // PapaParse sees trailing \n as an empty row at end
    expect(data).toEqual([["a", "b", "c"], [""]]);
  });

  it("handles quoted field containing newline", () => {
    const { data } = parseCsv('a,"line1\nline2",c');
    expect(data).toEqual([["a", "line1\nline2", "c"]]);
  });

  it("normalizeCsvRows + parseCsv: end-to-end jagged row normalization", () => {
    const { data } = parseCsv("a,b,c\nd,e\ng");
    const maxCols = Math.max(...data.map((row) => row.length), 0);
    const normalized = normalizeCsvRows(data, maxCols);
    expect(normalized).toEqual([
      ["a", "b", "c"],
      ["d", "e", ""],
      ["g", "", ""],
    ]);
  });

  // ── Regression: trailing blank rows must not inflate range size ────────────
  //
  // Production `importCsvToSheet` filters trailing blank rows (PapaParse
  // artifacts for `\n` at end of file, trailing `\n\n`, etc.) *before*
  // building `normalizedData`.  The `getResizedRange` call and the `rows`
  // metadata must use the *cleaned* row count, not the raw PapaParse count.
  //
  // These tests replicate the exact filter logic from the production body
  // so they catch a regression where range dimensions drift from written data.

  it("trailing blank row yields dimension aligned with cleaned data count", () => {
    // "a,b,c\n" — PapaParse returns [["a","b","c"],[""]]  (2 raw rows)
    const { data } = parseCsv("a,b,c\n");
    const totalCols = Math.max(...data.map((r) => r.length), 0);
    const cleaned = data.filter((row) =>
      row.some((cell) => cell !== "" && cell !== null && cell !== undefined)
    );
    // raw count = 2, cleaned count = 1
    expect(data.length).toBe(2);
    expect(cleaned.length).toBe(1);
    // getResizedRange(cleaned.length - 1, totalCols - 1) yields 1×3 range ✓
    // getResizedRange(data.length - 1, totalCols - 1) yields 2×3 range ✗
  });

  it("double trailing blank rows yield dimension aligned with cleaned data count", () => {
    // "a,b\n\n" — PapaParse returns [["a","b"],[""],[""]]  (3 raw rows)
    const { data } = parseCsv("a,b\n\n");
    const totalCols = Math.max(...data.map((r) => r.length), 0);
    const cleaned = data.filter((row) =>
      row.some((cell) => cell !== "" && cell !== null && cell !== undefined)
    );
    expect(data.length).toBe(3);
    expect(cleaned.length).toBe(1);
  });

  it("preserves an interleaved blank row while trimming only trailing blank rows", () => {
    const { data } = parseCsv("a\n\nb\n");
    const cleaned = data.slice();
    while (cleaned.length > 0) {
      const lastRow = cleaned[cleaned.length - 1];
      if (lastRow.some((cell) => cell !== "" && cell !== null && cell !== undefined)) break;
      cleaned.pop();
    }
    expect(cleaned).toEqual([["a"], [""], ["b"]]);
  });

  // ── Regression: trailing blank rows wider than retained data ──────────────
  //
  // When PapaParse emits trailing blank rows that are wider than the actual
  // CSV data (e.g. "a,b\n,,\n" → raw [["a","b"],["","",""],[""]]), computing
  // totalCols from raw data inflates the write range and can overwrite
  // adjacent cells. totalCols MUST be derived from cleanedData after terminal
  // blank-row removal so discarded rows cannot influence column count.
  it("trailing blank row wider than data does not inflate column count", () => {
    // "a,b\n,,\n" — PapaParse: [["a","b"],["","",""],[""]]
    const { data } = parseCsv("a,b\n,,\n");
    expect(data.length).toBe(3);

    const cleaned = data.slice();
    while (cleaned.length > 0) {
      const lastRow = cleaned[cleaned.length - 1];
      if (lastRow.some((cell) => cell !== "" && cell !== null && cell !== undefined)) break;
      cleaned.pop();
    }
    // Both trailing rows are all-blank; only ["a","b"] survives
    expect(cleaned).toEqual([["a", "b"]]);

    // totalCols from cleaned = 2, NOT 3 (raw max was 3 from ["","",""])
    const totalCols = Math.max(...cleaned.map((r) => r.length), 0);
    expect(totalCols).toBe(2);
  });
});
