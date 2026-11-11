<!-- SPDX-License-Identifier: Apache-2.0 -->

Hermes plugin tools use the `sidecar_excel_` prefix (for example `sidecar_excel_read_range`).
Unprefixed names below are the unchanged canonical Office.js/pane names.
See [desk-bridge.md](desk-bridge.md) for setup and security.

# Hermes Excel Tool Reference

This document lists the Office.js tools currently exposed by the Hermes Excel task pane add-in.

The command engine lives in:

```text
src/excelTools.ts
```

The plugin bundles `atlas-sidecar-excel:office-guidance` for ordinary spreadsheet, budget and chart work,
plus `atlas-sidecar-excel:workbook-plans` for plan mode. Each includes its portable support files;
no profile-installed `hermes-excel` operator guide is required.

## Current tools

| Tool | Purpose | Required args | Example |
|---|---|---|---|
| `get_workbook_overview` | Get sheet names, active sheet, used ranges, and current selection. | none | `{}` |
| `read_range` | Read values, formulas, and number formats. Truncates ranges above 5,000 cells. | `sheet`, `address` | `{ "sheet": "Sheet1", "address": "A1:D10" }` |
| `summarize_range_stats` | Compute deterministic range statistics: count, numericCount, sum, average, min, max, blankCount. | `sheet`, `address` | `{ "sheet": "Sheet1", "address": "D2:D20" }` |
| `write_range` | Write a 2D array to a range; strings beginning with `=` are formulas. | `sheet`, `address`, `values` | `{ "sheet": "Sheet1", "address": "E2:E3", "values": [["=SUM(A2:D2)"], ["=SUM(A3:D3)"]] }` |
| `format_range` | Apply font, fill, number format, borders, alignment, and optional autofit options. | `sheet`, `address`, `format` | `{ "sheet": "Sheet1", "address": "A1:D1", "format": { "bold": true, "fillColor": "#1f4e78" } }` |
| `conditional_format_range` | Apply common conditional formatting rules to a range (`cellValue`, `containsText`, and more in the future). | `sheet`, `address`, `ruleType` | `{ "sheet": "Sheet1", "address": "A1:A20", "ruleType": "cellValue", "operator": "GreaterThan", "formula1": "100", "fillColor": "#FFF2CC" }` |
| `freeze_panes` | Freeze rows and columns in a worksheet for easier navigation. | `sheet`, `rows`, `columns` | `{ "sheet": "Sheet1", "rows": 1, "columns": 0 }` |
| `set_data_validation` | Add list-style validation constraints (dropdowns) and optional prompt text. | `sheet`, `address`, `type`, `source` | `{ "sheet": "Sheet1", "address": "B2:B20", "type": "list", "source": "Yes,No,Maybe", "promptTitle": "Choose status", "promptMessage": "Pick one value." }` |
| `create_sheet` | Create and activate a worksheet. | `name` | `{ "name": "Analysis" }` |
| `delete_sheet` | Delete a worksheet. | `name` | `{ "name": "Old Sheet" }` |
| `activate_sheet` | Activate a worksheet. | `name` | `{ "name": "Sheet1" }` |
| `insert_table` | Create an Excel table from a range. | `sheet`, `address`, `hasHeaders` | `{ "sheet": "Sheet1", "address": "A1:D20", "hasHeaders": true }` |
| `sort_table` | Sort a table by zero-based column index. | `tableName`, `columnIndex` | `{ "tableName": "Table1", "columnIndex": 2, "ascending": false }` |
| `filter_table` | Filter a table column by values or custom criteria. | `tableName`, `columnIndex` | `{ "tableName": "Table1", "columnIndex": 2, "criterion1": ">1000" }` |
| `create_chart` | Create a chart next to a data range. | `sheet`, `dataRange`, `chartType`, `title` | `{ "sheet": "Sheet1", "dataRange": "A1:B10", "chartType": "ColumnClustered", "title": "Sales" }` |
| `find_replace` | Find and replace text in a sheet used range. | `sheet`, `find`, `replace`, `matchCase` | `{ "sheet": "Sheet1", "find": "old", "replace": "new", "matchCase": false }` |
| `clear_range` | Clear contents and formats from a range. | `sheet`, `address` | `{ "sheet": "Sheet1", "address": "A1:D10" }` |
| `run_calculation` | Recalculate the workbook. | none | `{}` |
| `autofill_formula` | Fill formulas from a single-cell source down or across a target range. Falls back to writing the formula directly if autoFill is unsupported. | `sheet`, `sourceAddress`, `targetAddress` | `{ "sheet": "Sheet1", "sourceAddress": "E2", "targetAddress": "E2:E100" }` |
| `rename_sheet` | Rename a worksheet. | `oldName`, `newName` | `{ "oldName": "Sheet1", "newName": "Data" }` |
| `copy_sheet` | Copy a worksheet to a new name within the same workbook. | `sourceName`, `newName` | `{ "sourceName": "Template", "newName": "January" }` |
| `protect_sheet` | Protect a worksheet. Password is never logged. | `sheet` | `{ "sheet": "Sheet1", "allowSort": true, "allowFilter": true }` |
| `unprotect_sheet` | Unprotect a worksheet. | `sheet` | `{ "sheet": "Sheet1", "password": "" }` |
| `create_pivot_table` | Create a pivot table from a source range with row fields, column fields, and aggregated data fields. | `sheet`, `address`, `sourceSheet`, `sourceRange`, `rowFields`, `dataFields` | `{ "sheet": "Pivot", "address": "A3", "sourceSheet": "Data", "sourceRange": "A1:D100", "rowFields": ["Region"], "dataFields": [{ "name": "Sales", "summarizeBy": "sum" }] }` |
| `add_comment` | Add or update a comment (note) on a specific cell. Returns clear error on hosts without Comments API (ExcelApi 1.10+). | `sheet`, `address`, `text` | `{ "sheet": "Sheet1", "address": "C5", "text": "Review this figure" }` |
| `clear_comments` | Delete all comments (notes) in a range. Detects host support and returns clear error on unsupported hosts. | `sheet`, `address` | `{ "sheet": "Sheet1", "address": "A1:C10" }` |
| `import_csv_to_sheet` | Parse a CSV string and write to a worksheet. Handles quoted commas, escaped quotes, CRLF, blank fields, and malformed data via PapaParse (RFC 4180). Delimiter auto-detection: `,` `\t` `;` `|`. Malformed CSV with structural errors throws a clear failure. | `sheet`, `address`, `csv` | `{ "sheet": "Sheet1", "address": "A1", "csv": "Name,Age\nAlice,30\nBob,25", "hasHeaders": true }` |

## Format options (format_range)

In addition to the standard formatting properties, `format_range` now supports:

| Property | Type | Description |
|---|---|---|
| `fontName` | string | Font family (e.g. "Calibri", "Arial") |
| `fontSize` | number | Font size in points |
| `underline` | "None" \| "Single" \| "Double" | Underline style |
| `columnAutofit` | boolean | Auto-fit columns when `true`. Defaults to `false` (explicit opt-in). |
| `rowAutofit` | boolean | Auto-fit rows when `true`. Defaults to `false` (explicit opt-in). |
| `wrapText` | boolean | Wrap text within cells |
| `indentLevel` | integer | Indentation level |
| `textOrientation` | integer | Rotation angle in degrees (-90 to 90) |
| `borderColor` | string | Border color hex (overrides default #c7cdd1 when borders=true) |
| `borderWeight` | "Hairline" \| "Thin" \| "Medium" \| "Thick" | Border weight (default Thin) |

### Horizontal alignment (horizontalAlignment)

Canonical accepted values: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterAcrossSelection`, `Distributed`.

`CenterAcrossSelection` spans a title across columns without merging cells. It is safe for print/screen-reader navigation but requires preconditions:

1. The span must be a **single row** with **at least two columns**.
2. Every non-anchor cell must be **blank** (no value or formula).
3. The span must not be **already merged**.

If preconditions fail, the tool returns a structured error and **no formatting is applied**. No merge or unmerge capability exists; refuse and explain.

### Title-span recipe (prompt guidance)

```
For a title spanning columns, first read the full intended single-row span.
Proceed only when every non-anchor cell is blank and the span is not merged.
Write the title only to the leftmost anchor, then format the full span with
horizontalAlignment CenterAcrossSelection, wrapText false, and columnAutofit
false. No merge capability exists; refuse merge requests and explain.
```

## Safe usage patterns

1. Call `get_workbook_overview` before acting on vague references like “this sheet” or “the selected range.”
2. Call `read_range` before changing existing data.
3. Use exact sheet names returned by workbook context.
4. Keep edits small and verifiable.
5. After writing formulas, call `run_calculation`.
6. Confirm destructive actions if the user’s scope is ambiguous.
7. The same exact tool call (same name and full arguments JSON) is blocked if repeated more than two times in one run; treat the returned error as guidance to reuse prior result or ask for clarification.

## Completed roadmap items

All items from the initial roadmap are now implemented:

- `create_pivot_table` — pivot table creation with row/column/data fields (require ExcelApi 1.8+)
- `autofill_formula` — formula fill with fallback
- `rename_sheet` / `copy_sheet` — workbook organization
- `protect_sheet` / `unprotect_sheet` — sheet security with optional password
- `format_range` expansion — fontName, fontSize, underline, wrapText, indentLevel, textOrientation, borderColor, borderWeight
- `add_comment` — cell comment/notes (ExcelApi 1.10+)
- `clear_comments` — cell comment removal
- `import_csv_to_sheet` — RFC 4180 CSV parsing via PapaParse with header formatting

## Capability matrix

See [`docs/capability-matrix.md`](capability-matrix.md) for the full capability status across all categories.

## Validation

Run from the repo root:

```bash
npm run test
npm run build:static -- --base https://example.org/excel/
hermes plugins validate plugins/atlas-sidecar-excel
```
