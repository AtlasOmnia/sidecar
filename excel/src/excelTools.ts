import { OpenAITool, ToolCall } from "./types";
import * as Papa from "papaparse";

type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue };

export interface RangeStats {
  count: number;
  numericCount: number;
  sum: number;
  average: number | null;
  min: number | null;
  max: number | null;
  blankCount: number;
}

const MAX_READ_CELLS = 5000;

export interface ExcelToolDefinition {
  schema: OpenAITool;
  execute: (args: Record<string, unknown>) => Promise<string>;
}

type ConditionalFormatRuleType = "cellValue" | "containsText";
type SupportedConditionalFormatRuleType = "cellValue" | "containsText";

type DataValidationType = "list";

export const excelTools: ExcelToolDefinition[] = [
  {
    schema: {
      type: "function",
      function: {
        name: "get_workbook_overview",
        description: "Get workbook sheets, active sheet, used ranges, and current selection.",
        parameters: { type: "object", properties: {}, additionalProperties: false }
      }
    },
    execute: () => getWorkbookOverview()
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_range",
        description: "Read values, formulas, and number formats from a range. Results larger than 5000 cells are truncated.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string", description: "A1 address such as A1:C10." }
          },
          required: ["sheet", "address"],
          additionalProperties: false
        }
      }
  },
    execute: (args) => readRange(textArg(args, "sheet"), textArg(args, "address"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "summarize_range_stats",
        description: "Compute deterministic numeric statistics for a range: count, numericCount, sum, average, min, max, and blankCount.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string", description: "A1 address such as A1:C10." }
          },
          required: ["sheet", "address"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => summarizeRangeStats(textArg(args, "sheet"), textArg(args, "address"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "write_range",
        description: "Write a 2D array to a range. Strings beginning with = are written as formulas.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
            values: { type: "array", items: { type: "array" } }
          },
          required: ["sheet", "address", "values"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => writeRange(textArg(args, "sheet"), textArg(args, "address"), array2dArg(args, "values"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "format_range",
        description: "Format a range with font, fill, number format, borders, alignment, and autofit options. Use CenterAcrossSelection to span a title across columns without merging.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
            format: {
              type: "object",
              properties: {
                bold: { type: "boolean" },
                italic: { type: "boolean" },
                fontColor: { type: "string" },
                fillColor: { type: "string" },
                numberFormat: { type: "string" },
                borders: { type: "boolean" },
                columnAutofit: { type: "boolean" },
                rowAutofit: { type: "boolean" },
                horizontalAlignment: { type: "string", enum: ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"], description: "Horizontal alignment. CenterAcrossSelection formats a single-row multi-cell span without merging." },
                verticalAlignment: { type: "string" },
                fontName: { type: "string", description: "Font family name such as Calibri or Arial." },
                fontSize: { type: "number", description: "Font size in points." },
                underline: { type: "string", enum: ["None", "Single", "Double"], description: "Underline style." },
                wrapText: { type: "boolean", description: "Whether to wrap text in the range." },
                indentLevel: { type: "integer", description: "Indentation level (integer)." },
                textOrientation: { type: "integer", description: "Text orientation in degrees, typically -90 to 90." },
                borderColor: { type: "string", description: "Border color hex (overrides default border color when borders=true)." },
                borderWeight: { type: "string", enum: ["Hairline", "Thin", "Medium", "Thick"], description: "Border weight when borders=true." }
              },
              additionalProperties: false
            }
          },
          required: ["sheet", "address", "format"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => formatRange(textArg(args, "sheet"), textArg(args, "address"), objectArg(args, "format"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "conditional_format_range",
        description: "Apply conditional formatting to a range using common rule types.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
            ruleType: {
              type: "string",
              enum: ["cellValue", "containsText"]
            },
            operator: { type: "string" },
            formula1: { type: "string" },
            formula2: { type: "string" },
            text: { type: "string" },
            fillColor: { type: "string" },
            fontColor: { type: "string" }
          },
          required: ["sheet", "address", "ruleType"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => {
      try {
        const ruleType = parseConditionalFormatRuleType(textArg(args, "ruleType"));
        return conditionalFormatRange(textArg(args, "sheet"), textArg(args, "address"), ruleType, args);
      } catch (error) {
        return Promise.resolve(stringifyError(error));
      }
    }
  },
  {
    schema: {
      type: "function",
      function: {
        name: "freeze_panes",
        description: "Freeze rows or columns in a worksheet.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            rows: { type: "number" },
            columns: { type: "number" }
          },
          required: ["sheet", "rows", "columns"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => freezePanes(
      textArg(args, "sheet"),
      nonNegativeIntArg(args, "rows"),
      nonNegativeIntArg(args, "columns")
    )
  },
  {
    schema: {
      type: "function",
      function: {
        name: "set_data_validation",
        description: "Apply data validation rules to a range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
            type: { type: "string", enum: ["list"] },
            source: { type: "string", description: "Comma-separated list values, for example \"Yes,No,Maybe\"." },
            promptTitle: { type: "string" },
            promptMessage: { type: "string" }
          },
          required: ["sheet", "address", "type", "source"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => setDataValidation(
      textArg(args, "sheet"),
      textArg(args, "address"),
      textArg(args, "type"),
      textArg(args, "source"),
      optionalTextArg(args, "promptTitle"),
      optionalTextArg(args, "promptMessage")
    )
  },
  sheetTool("create_sheet", "Create a worksheet.", ["name"], (args) => createSheet(textArg(args, "name"))),
  sheetTool("delete_sheet", "Delete a worksheet.", ["name"], (args) => deleteSheet(textArg(args, "name"))),
  sheetTool("activate_sheet", "Activate a worksheet.", ["name"], (args) => activateSheet(textArg(args, "name"))),
  {
    schema: {
      type: "function",
      function: {
        name: "insert_table",
        description: "Create an Excel table from a range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
            hasHeaders: { type: "boolean" }
          },
          required: ["sheet", "address", "hasHeaders"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => insertTable(textArg(args, "sheet"), textArg(args, "address"), boolArg(args, "hasHeaders"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "sort_table",
        description: "Sort a table by zero-based column index.",
        parameters: {
          type: "object",
          properties: {
            tableName: { type: "string" },
            columnIndex: { type: "number" },
            ascending: { type: "boolean" }
          },
          required: ["tableName", "columnIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => sortTable(textArg(args, "tableName"), nonNegativeIntArg(args, "columnIndex"), args.ascending !== false)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "filter_table",
        description: "Filter a table column by values or a custom criterion such as >1000.",
        parameters: {
          type: "object",
          properties: {
            tableName: { type: "string" },
            columnIndex: { type: "number" },
            values: { type: "array", items: { type: "string" } },
            criterion1: { type: "string" },
            criterion2: { type: "string" },
            operator: { type: "string", enum: ["And", "Or"] }
          },
          required: ["tableName", "columnIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => filterTable(textArg(args, "tableName"), nonNegativeIntArg(args, "columnIndex"), args)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "create_chart",
        description: "Create a chart next to a data range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            dataRange: { type: "string" },
            chartType: { type: "string", description: "Excel chart type such as ColumnClustered, Line, Pie, BarClustered." },
            title: { type: "string" }
          },
          required: ["sheet", "dataRange", "chartType", "title"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => createChart(textArg(args, "sheet"), textArg(args, "dataRange"), textArg(args, "chartType"), textArg(args, "title"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "find_replace",
        description: "Find and replace text within the used range of a sheet.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            find: { type: "string" },
            replace: { type: "string" },
            matchCase: { type: "boolean" }
          },
          required: ["sheet", "find", "replace", "matchCase"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => findReplace(textArg(args, "sheet"), textArg(args, "find"), textArg(args, "replace"), boolArg(args, "matchCase"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "clear_range",
        description: "Clear all contents and formats from a range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string" }
          },
          required: ["sheet", "address"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => clearRange(textArg(args, "sheet"), textArg(args, "address"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "autofill_formula",
        description: "Fill formulas from a source cell down or across a target range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            sourceAddress: { type: "string", description: "Single-cell address with the formula to copy, such as E2." },
            targetAddress: { type: "string", description: "Target range such as E2:E100." }
          },
          required: ["sheet", "sourceAddress", "targetAddress"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => autofillFormula(textArg(args, "sheet"), textArg(args, "sourceAddress"), textArg(args, "targetAddress"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "rename_sheet",
        description: "Rename a worksheet.",
        parameters: {
          type: "object",
          properties: {
            oldName: { type: "string" },
            newName: { type: "string" }
          },
          required: ["oldName", "newName"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => renameSheet(textArg(args, "oldName"), textArg(args, "newName"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "copy_sheet",
        description: "Copy a worksheet to a new name within the same workbook.",
        parameters: {
          type: "object",
          properties: {
            sourceName: { type: "string" },
            newName: { type: "string" }
          },
          required: ["sourceName", "newName"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => copySheet(textArg(args, "sourceName"), textArg(args, "newName"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "protect_sheet",
        description: "Protect a worksheet. Password is never logged or stored; if the user did not provide a password, protect without one.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            password: { type: "string", description: "Optional password. If omitted, protect without a password." },
            allowSort: { type: "boolean" },
            allowFilter: { type: "boolean" }
          },
          required: ["sheet"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => protectSheet(textArg(args, "sheet"), optionalTextArg(args, "password"), optionalBoolArg(args, "allowSort"), optionalBoolArg(args, "allowFilter"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "unprotect_sheet",
        description: "Unprotect a worksheet.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            password: { type: "string", description: "Optional password. Required only if the sheet was protected with one." }
          },
          required: ["sheet"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => unprotectSheet(textArg(args, "sheet"), optionalTextArg(args, "password"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "create_pivot_table",
        description: "Create a pivot table from a source range. Returns an informative error if the Excel version does not support pivot creation through Office.js.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string", description: "Target sheet name where the pivot table is created." },
            address: { type: "string", description: "Top-left cell address for the pivot table output, such as A3." },
            sourceSheet: { type: "string", description: "Source sheet name." },
            sourceRange: { type: "string", description: "Source data A1 address such as A1:D100. Must include headers." },
            rowFields: { type: "array", items: { type: "string" }, description: "Column names to use as row fields." },
            columnFields: { type: "array", items: { type: "string" }, description: "Column names to use as column fields." },
            dataFields: { type: "array", items: { type: "object", properties: { name: { type: "string" }, summarizeBy: { type: "string", enum: ["sum", "count", "average", "max", "min"] } }, required: ["name"] }, description: "Data fields to aggregate." }
          },
          required: ["sheet", "address", "sourceSheet", "sourceRange", "rowFields", "dataFields"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => createPivotTable(
      textArg(args, "sheet"),
      textArg(args, "address"),
      textArg(args, "sourceSheet"),
      textArg(args, "sourceRange"),
      args.rowFields as string[] | undefined,
      args.columnFields as string[] | undefined,
      args.dataFields as Array<Record<string, string>> | undefined
    )
  },
  {
    schema: {
      type: "function",
      function: {
        name: "run_calculation",
        description: "Recalculate the workbook.",
        parameters: { type: "object", properties: {}, additionalProperties: false }
      }
    },
    execute: () => runCalculation()
  },
  {
    schema: {
      type: "function",
      function: {
        name: "add_comment",
        description: "Add or update a comment (note) on a specific cell.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string", description: "Single-cell A1 address such as C5." },
            text: { type: "string", description: "Comment text content." }
          },
          required: ["sheet", "address", "text"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => addComment(textArg(args, "sheet"), textArg(args, "address"), textArg(args, "text"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "clear_comments",
        description: "Delete all comments (notes) in a range.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string", description: "A1 address of the range to clear comments from." }
          },
          required: ["sheet", "address"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => clearComments(textArg(args, "sheet"), textArg(args, "address"))
  },
  {
    schema: {
      type: "function",
      function: {
        name: "import_csv_to_sheet",
        description: "Parse a CSV string and write the resulting 2D data array into a worksheet starting at the given address.",
        parameters: {
          type: "object",
          properties: {
            sheet: { type: "string" },
            address: { type: "string", description: "Top-left cell address to place the CSV data, such as A1." },
            csv: { type: "string", description: "CSV data as a plain-text string. Delimiter auto-detection tries comma, tab, semicolon, pipe in order. The first row is treated as headers (bolded) by default when hasHeaders is true." },
            hasHeaders: { type: "boolean", description: "Whether the first CSV row contains column headers (default true). If true, the header row is bolded. Note: delimiter auto-detection tries comma, tab, semicolon, pipe in order." }
          },
          required: ["sheet", "address", "csv"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => importCsvToSheet(textArg(args, "sheet"), textArg(args, "address"), textArg(args, "csv"), optionalBoolArg(args, "hasHeaders") ?? true)
  }
];

function buildExcelToolDefinitionLookup(
  tools: readonly ExcelToolDefinition[],
): Map<string, ExcelToolDefinition> {
  const lookup = new Map<string, ExcelToolDefinition>();

  for (const tool of tools) {
    const name = tool.schema.function.name;
    if (lookup.has(name)) {
      throw new Error("Duplicate Excel tool primitive names detected.");
    }
    lookup.set(name, tool);
  }

  return lookup;
}

const excelToolDefinitionLookup = buildExcelToolDefinitionLookup(excelTools);

export function getExcelToolDefinition(name: string): ExcelToolDefinition | undefined {
  return excelToolDefinitionLookup.get(name);
}

export const toolSchemas = excelTools.map((tool) => tool.schema);

/*
 * R4a: effect class of every Excel primitive (D-R4-1 (a), D-R4-2).
 * - read_only: never mutates the workbook.
 * - mutating: a workbook-plan compiler target; planEffect/planKinds mirror
 *   workbookPlanContract.operationDefinitions (pinned by tests, not imported,
 *   so this module stays standalone for scripts/generate-office-plugin-schemas.mjs).
 * - view: changes only what the user sees (activate_sheet).
 * - destructive_legacy: no plan equivalent; to be failed closed while plan rollout is on (R4c/R4d).
 * Consumed by src/primitiveMutationAuthority.ts (R4c) and the read-only sinks below. Completeness is enforced by
 * src/__tests__/primitive-effect-table.test.ts; a consumer must treat a missing entry as deny.
 */
export type PrimitiveEffectClass = "read_only" | "mutating" | "view" | "destructive_legacy";
export type PrimitivePlanEffect = "write" | "format" | "structure" | "filter" | "calculate";

export interface PrimitiveEffect {
  readonly effectClass: PrimitiveEffectClass;
  readonly planEffect?: PrimitivePlanEffect;
  readonly planKinds?: readonly string[];
}

export const PRIMITIVE_EFFECTS: Readonly<Record<string, PrimitiveEffect>> = Object.freeze({
  get_workbook_overview: Object.freeze({ effectClass: "read_only" }),
  read_range: Object.freeze({ effectClass: "read_only" }),
  summarize_range_stats: Object.freeze({ effectClass: "read_only" }),
  write_range: Object.freeze({ effectClass: "mutating", planEffect: "write", planKinds: Object.freeze(["write_values", "write_formula"]) }),
  autofill_formula: Object.freeze({ effectClass: "mutating", planEffect: "write", planKinds: Object.freeze(["fill_formula"]) }),
  format_range: Object.freeze({ effectClass: "mutating", planEffect: "format", planKinds: Object.freeze(["format_range"]) }),
  conditional_format_range: Object.freeze({ effectClass: "mutating", planEffect: "format", planKinds: Object.freeze(["conditional_format"]) }),
  freeze_panes: Object.freeze({ effectClass: "mutating", planEffect: "format", planKinds: Object.freeze(["freeze_panes"]) }),
  set_data_validation: Object.freeze({ effectClass: "mutating", planEffect: "write", planKinds: Object.freeze(["set_validation_list"]) }),
  insert_table: Object.freeze({ effectClass: "mutating", planEffect: "structure", planKinds: Object.freeze(["create_table"]) }),
  sort_table: Object.freeze({ effectClass: "mutating", planEffect: "write", planKinds: Object.freeze(["sort_table"]) }),
  filter_table: Object.freeze({ effectClass: "mutating", planEffect: "filter", planKinds: Object.freeze(["filter_table", "reset_table_filter"]) }),
  create_chart: Object.freeze({ effectClass: "mutating", planEffect: "structure", planKinds: Object.freeze(["create_chart"]) }),
  create_sheet: Object.freeze({ effectClass: "mutating", planEffect: "structure", planKinds: Object.freeze(["create_sheet"]) }),
  rename_sheet: Object.freeze({ effectClass: "mutating", planEffect: "structure", planKinds: Object.freeze(["rename_sheet"]) }),
  copy_sheet: Object.freeze({ effectClass: "mutating", planEffect: "structure", planKinds: Object.freeze(["copy_sheet"]) }),
  add_comment: Object.freeze({ effectClass: "mutating", planEffect: "write", planKinds: Object.freeze(["add_comment"]) }),
  run_calculation: Object.freeze({ effectClass: "mutating", planEffect: "calculate", planKinds: Object.freeze(["calculate"]) }),
  activate_sheet: Object.freeze({ effectClass: "view" }),
  delete_sheet: Object.freeze({ effectClass: "destructive_legacy" }),
  clear_range: Object.freeze({ effectClass: "destructive_legacy" }),
  clear_comments: Object.freeze({ effectClass: "destructive_legacy" }),
  find_replace: Object.freeze({ effectClass: "destructive_legacy" }),
  unprotect_sheet: Object.freeze({ effectClass: "destructive_legacy" }),
  protect_sheet: Object.freeze({ effectClass: "destructive_legacy" }),
  import_csv_to_sheet: Object.freeze({ effectClass: "destructive_legacy" }),
  create_pivot_table: Object.freeze({ effectClass: "destructive_legacy" }),
});

export function getPrimitiveEffect(name: string): PrimitiveEffect | undefined {
  if (typeof name !== "string" || !Object.prototype.hasOwnProperty.call(PRIMITIVE_EFFECTS, name)) {
    return undefined;
  }
  return PRIMITIVE_EFFECTS[name];
}

export function missingPrimitiveEffects(names: readonly string[]): string[] {
  return names.filter((name) => getPrimitiveEffect(name) === undefined);
}

/**
 * R4c: fixed denial for a non-read-only primitive reaching an exported sink.
 * Workbook edits are executed only by the per-turn task-pane authority
 * (src/primitiveMutationAuthority.ts), which holds the raw dispatcher.
 */
const READ_ONLY_SINK_RESULT = JSON.stringify({
  ok: false,
  error: "Workbook edits must go through the task-pane edit authority.",
  category: "validation",
});

let rawExecutorClaimed = false;

/**
 * Raw name -> tool.execute dispatcher. Module-private; handed out exactly once
 * through claimRawPrimitiveExecutor().
 */
async function executeRawPrimitive(
  name: string,
  args: Record<string, unknown>,
): Promise<string> {
  const tool = getExcelToolDefinition(name);
  if (!tool) {
    return stringifyError(`Unknown tool: ${name}`);
  }

  try {
    return await tool.execute(args);
  } catch (error) {
    return stringifyError(error);
  }
}

/** One-shot latch: the raw dispatcher is claimed once, by src/primitiveMutationAuthority.ts. */
export function claimRawPrimitiveExecutor(): (name: string, args: Record<string, unknown>) => Promise<string> {
  if (rawExecutorClaimed) {
    throw new Error("Primitive executor already claimed.");
  }
  rawExecutorClaimed = true;
  return executeRawPrimitive;
}

export async function executeExcelTool(name: string, rawArgs: string): Promise<string> {
  const tool = getExcelToolDefinition(name);
  if (!tool) {
    return stringifyError(`Unknown tool: ${name}`);
  }
  if (getPrimitiveEffect(name)?.effectClass !== "read_only") {
    return READ_ONLY_SINK_RESULT;
  }

  try {
    const args = rawArgs.trim() ? JSON.parse(rawArgs) as Record<string, unknown> : {};
    return await tool.execute(args);
  } catch (error) {
    return stringifyError(error);
  }
}

/**
 * Execute a tool with already-parsed and validated arguments.
 * This is the canonical execution path — callers that have already
 * validated via validateToolCall() should use this to avoid re-parsing
 * (repair 1).
 */
export async function executeExcelToolParsed(
  name: string,
  args: Record<string, unknown>,
): Promise<string> {
  const tool = getExcelToolDefinition(name);
  if (!tool) {
    return stringifyError(`Unknown tool: ${name}`);
  }
  if (getPrimitiveEffect(name)?.effectClass !== "read_only") {
    return READ_ONLY_SINK_RESULT;
  }

  return executeRawPrimitive(name, args);
}

export async function getWorkbookContext(): Promise<string> {
  return getWorkbookOverview();
}

async function getWorkbookOverview(): Promise<string> {
  return safelyExcel(async (context) => {
    const workbook = context.workbook;
    const sheets = workbook.worksheets;
    const activeSheet = sheets.getActiveWorksheet();
    const selection = workbook.getSelectedRange();

    // Sync 1 — hydrate sheet names, active sheet, and selection address
    sheets.load("items/name");
    activeSheet.load("name");
    selection.load("address");
    await context.sync();

    // Now sheets.items is populated; queue used-range loads
    const usedRanges = sheets.items.map((sheet) => {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["address", "rowCount", "columnCount"]);
      return { sheet, usedRange };
    });
    await context.sync();

    const sheetOverviews = usedRanges.map(({ sheet, usedRange }) => ({
      name: sheet.name,
      usedRange: usedRange.isNullObject ? null : {
        address: stripSheetName(usedRange.address),
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount
      }
    }));

    return {
      activeSheet: activeSheet.name,
      selection: selection.address,
      sheets: sheetOverviews
    };
  });
}

async function readRange(sheetName: string, address: string): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    const localDimensions = parseA1RangeDimensions(address);
    const deterministicAddress = `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`;
    const localCellCount = localDimensions
      ? localDimensions.rows * localDimensions.columns
      : null;

    if (localDimensions && localCellCount !== null && localCellCount > MAX_READ_CELLS) {
      const previewRows = Math.max(1, Math.floor(MAX_READ_CELLS / localDimensions.columns));
      const preview = range.getCell(0, 0).getResizedRange(previewRows - 1, localDimensions.columns - 1);
      const deterministicPreviewAddress = getPreviewA1Address(sheetName, address, previewRows, localDimensions.columns);
      if (deterministicPreviewAddress === null) {
        preview.load(["address", "values", "formulas", "numberFormat"]);
      } else {
        preview.load(["values", "formulas", "numberFormat"]);
      }
      await context.sync();

      return {
        sheet: sheetName,
        address: deterministicAddress,
        rowCount: localDimensions.rows,
        columnCount: localDimensions.columns,
        truncated: true,
        note: `Range has ${localCellCount} cells. Returning the first ${previewRows} rows only.`,
        previewAddress: deterministicPreviewAddress ?? preview.address,
        values: preview.values,
        formulas: preview.formulas,
        numberFormat: preview.numberFormat
      };
    }

    range.load(localDimensions
      ? ["rowCount", "columnCount", "values", "formulas", "numberFormat"]
      : ["address", "rowCount", "columnCount", "values", "formulas", "numberFormat"]);
    await context.sync();

    const cellCount = range.rowCount * range.columnCount;
    if (cellCount <= MAX_READ_CELLS) {
      return {
        sheet: sheetName,
        address: localDimensions ? deterministicAddress : range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat
      };
    }

    const previewRows = Math.max(1, Math.floor(MAX_READ_CELLS / Math.max(range.columnCount, 1)));
    const preview = range.getCell(0, 0).getResizedRange(previewRows - 1, range.columnCount - 1);
    preview.load(["address", "values", "formulas", "numberFormat"]);
    await context.sync();

    return {
      sheet: sheetName,
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      truncated: true,
      note: `Range has ${cellCount} cells. Returning the first ${previewRows} rows only.`,
      previewAddress: preview.address,
      values: preview.values,
      formulas: preview.formulas,
      numberFormat: preview.numberFormat
    };
  });
}

async function summarizeRangeStats(sheetName: string, address: string): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    const deterministicAddress = parseA1RangeDimensions(address)
      ? `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`
      : null;
    range.load(deterministicAddress === null ? ["address", "values"] : "values");
    await context.sync();

    const stats = summarizeRangeValues(range.values);
    return {
      action: "summarize_range_stats",
      sheet: sheetName,
      address: deterministicAddress ?? range.address,
      ...stats
    };
  });
}

async function writeRange(sheetName: string, address: string, values: JsonValue[][]): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    const knownDimensions = parseA1RangeDimensions(address);

    if (knownDimensions) {
      // Known dimensions from a strict A1 address let us fail fast before queuing
      // potentially destructive mutations.
      assertDimensions(values, knownDimensions.rows, knownDimensions.columns);
      range.formulas = values as Excel.Range["formulas"];
      range.select();
      await context.sync();

      const quotedSheetName = formatSheetNameForFormula(sheetName);
      const requestedRange = stripSheetName(address).replace(/\$/g, "").trim();
      return {
        ok: true,
        action: "write_range",
        sheet: sheetName,
        address: `${quotedSheetName}!${requestedRange}`,
        rows: knownDimensions.rows,
        columns: knownDimensions.columns
      };
    }

    range.load(["address", "rowCount", "columnCount"]);
    await context.sync();

    // Validate dimensions before mutating non-strict addresses, preserving deterministic failure
    // behavior rather than queuing an invalid write before host sync.
    assertDimensions(values, range.rowCount, range.columnCount);

    range.formulas = values as Excel.Range["formulas"];
    range.select();
    await context.sync();

    return { ok: true, action: "write_range", sheet: sheetName, address: range.address, rows: range.rowCount, columns: range.columnCount };
  });
}

/**
 * Canonical alignment values that Office.js accepts.
 * Specifically: CenterAcrossSelection spans a single-row multi-cell span without merging.
 * Runtime accepts case-insensitive legacy aliases (e.g. "centerAcrossSelection") and returns
 * the PascalCase canonical value.
 */
export const CANONICAL_HORIZONTAL_ALIGNMENT = [
  "General", "Left", "Center", "Right", "Fill", "Justify",
  "CenterAcrossSelection", "Distributed",
] as const;
export type CanonicalHorizontalAlignment = (typeof CANONICAL_HORIZONTAL_ALIGNMENT)[number];

// Build a lookup map for case-insensitive legacy-alias resolution.
const _alignmentAliasMap: Record<string, CanonicalHorizontalAlignment> = {};
for (const canonical of CANONICAL_HORIZONTAL_ALIGNMENT) {
  _alignmentAliasMap[canonical.toLowerCase()] = canonical;
}

/**
 * Normalize an alignment string to a canonical PascalCase enum value.
 * Accepts case-insensitive legacy aliases (e.g. "centerAcrossSelection" → "CenterAcrossSelection").
 * Throws on unrecognized alignment values.
 */
export function normalizeHorizontalAlignment(raw: unknown): CanonicalHorizontalAlignment {
  if (typeof raw !== "string") {
    throw new Error("horizontalAlignment must be a string.");
  }
  if (!raw.trim()) {
    throw new Error("horizontalAlignment must not be empty.");
  }
  const key = raw.trim().toLowerCase();
  const canonical = _alignmentAliasMap[key];
  if (!canonical) {
    throw new Error(`Unrecognized horizontalAlignment "${raw}". Use one of: ${CANONICAL_HORIZONTAL_ALIGNMENT.join(", ")}.`);
  }
  return canonical;
}

/**
 * For CenterAcrossSelection: validate that the span is a single row, at least
 * two columns wide, has no non-anchor values, and is not merged.
 * Returns an error message string on failure, or null if OK.
 *
 * This is a synchronous helper because the Office.js loads and sync must
 * happen in the calling context (within safelyExcel).
 */
export function validateCenterAcrossSelectionPrecondition(params: {
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: unknown[][];
  mergedAreasOrNullObject?: { isNullObject: boolean };
}): string | null {
  const { rowCount, columnCount, values, formulas, mergedAreasOrNullObject } = params;

  if (rowCount !== 1) {
    return "CenterAcrossSelection requires a single-row span.";
  }
  if (columnCount < 2) {
    return "CenterAcrossSelection requires at least two columns.";
  }
  if (mergedAreasOrNullObject?.isNullObject === false) {
    return "CenterAcrossSelection cannot be applied to a span that is already merged.";
  }
  // Check non-anchor cells (columns 1..n) for values or formulas.
  for (let col = 1; col < columnCount; col++) {
    const val = values[0]?.[col];
    const formula = formulas[0]?.[col];
    // null/undefined/empty string is OK for a non-anchor cell.
    if (val !== null && val !== undefined && val !== "") {
      return `CenterAcrossSelection failed: non-anchor cell ${String.fromCharCode(65 + col)}1 contains "${val}". The span must be blank except for the title anchor in the first cell.`;
    }
    // If the formula is non-empty, it's a problem too.
    if (typeof formula === "string" && formula !== "") {
      return `CenterAcrossSelection failed: non-anchor cell ${String.fromCharCode(65 + col)}1 contains a formula "${formula}". The span must be blank except for the title anchor in the first cell.`;
    }
  }
  return null;
}

async function formatRange(sheetName: string, address: string, format: Record<string, unknown>): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    const parsedDimensions = parseA1RangeDimensions(address);
    const localDimensions = typeof format.numberFormat === "string" ? parsedDimensions : null;
    const deterministicAddress = parsedDimensions === null
      ? null
      : `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`;
    const needsAddressLoad = deterministicAddress === null;
    const needsDimensionedNumberFormat = typeof format.numberFormat === "string" && localDimensions === null;
    const normalizedAlignment =
      typeof format.horizontalAlignment === "string" ? normalizeHorizontalAlignment(format.horizontalAlignment) : null;

    if (normalizedAlignment === "CenterAcrossSelection") {
      const centerAcrossSupported = isExcelApiSetSupported("ExcelApi", "1.13");
      if (!centerAcrossSupported) {
        if (needsAddressLoad) {
          range.load("address");
          await context.sync();
        }
        return {
          ok: false,
          error: "CenterAcrossSelection requires ExcelApi 1.13 and is not available in this Excel host. No formatting changes were applied.",
          action: "format_range",
          sheet: sheetName,
          address: deterministicAddress ?? range.address
        };
      }

      // Load all CAS preconditions before mutating anything in the same batch.
      range.load(["rowCount", "columnCount", "values", "formulas", ...(needsAddressLoad ? ["address"] : [])]);
      const merged = range.getMergedAreasOrNullObject();
      merged.load("isNullObject");
      await context.sync();

      const preconditionError = validateCenterAcrossSelectionPrecondition({
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values as unknown[][],
        formulas: range.formulas as unknown[][],
        mergedAreasOrNullObject: merged as { isNullObject: boolean },
      });
      if (preconditionError !== null) {
        return {
          ok: false,
          error: preconditionError,
          action: "format_range",
          sheet: sheetName,
          address: deterministicAddress ?? range.address
        };
      }
    } else if (needsDimensionedNumberFormat) {
      const fieldsToLoad = needsAddressLoad ? ["address", "rowCount", "columnCount"] : ["rowCount", "columnCount"];
      range.load(fieldsToLoad);
      await context.sync();
    } else if (needsAddressLoad) {
      range.load("address");
    }

    // Apply formatting mutations.
    if (typeof format.bold === "boolean") {
      range.format.font.bold = format.bold;
    }
    if (typeof format.italic === "boolean") {
      range.format.font.italic = format.italic;
    }
    if (typeof format.fontColor === "string") {
      range.format.font.color = format.fontColor;
    }
    if (typeof format.fillColor === "string") {
      range.format.fill.color = format.fillColor;
    }
    if (typeof format.numberFormat === "string") {
      const dimensions = localDimensions ?? { rows: range.rowCount, columns: range.columnCount };
      range.numberFormat = Array.from({ length: dimensions.rows }, () =>
        Array.from({ length: dimensions.columns }, () => format.numberFormat as string)
      );
    }

    if (normalizedAlignment !== null) {
      range.format.horizontalAlignment = normalizedAlignment as Excel.HorizontalAlignment;
    }

    if (typeof format.verticalAlignment === "string") {
      range.format.verticalAlignment = format.verticalAlignment as Excel.VerticalAlignment;
    }
    if (format.borders === true) {
      const borderIds = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
      const borderColor = typeof format.borderColor === "string" ? format.borderColor : "#c7cdd1";
      const borderWeight = typeof format.borderWeight === "string" ? format.borderWeight as Excel.BorderWeight : "Thin";
      for (const borderId of borderIds) {
        const border = range.format.borders.getItem(borderId as Excel.BorderIndex);
        border.style = "Continuous";
        border.color = borderColor;
        border.weight = borderWeight;
      }
    }
    if (typeof format.fontName === "string") {
      range.format.font.name = format.fontName;
    }
    if (typeof format.fontSize === "number") {
      range.format.font.size = format.fontSize;
    }
    if (typeof format.underline === "string") {
      range.format.font.underline = format.underline as Excel.RangeUnderlineStyle;
    }
    if (typeof format.wrapText === "boolean") {
      range.format.wrapText = format.wrapText;
    }
    if (typeof format.indentLevel === "number") {
      range.format.indentLevel = format.indentLevel;
    }
    if (typeof format.textOrientation === "number") {
      range.format.textOrientation = format.textOrientation;
    }
    // Omitted/missing columnAutofit and rowAutofit now explicit opt-in (=== true).
    if (format.columnAutofit === true) {
      range.format.autofitColumns();
    }
    if (format.rowAutofit === true) {
      range.format.autofitRows();
    }

    range.select();
    await context.sync();

    return { ok: true, action: "format_range", sheet: sheetName, address: deterministicAddress ?? range.address };
  });
}

type ExcelApiRequirements = {
  context?: {
    requirements?: {
      isSetSupported?: (apiSet: string, version: string) => boolean;
    };
  };
};

function isExcelApiSetSupported(apiSet: string, version: string): boolean {
  const office = (globalThis as { Office?: ExcelApiRequirements }).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") {
    return false;
  }

  try {
    return isSetSupported(apiSet, version);
  } catch {
    return false;
  }
}

async function createSheet(name: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = context.workbook.worksheets.add(name);
    sheet.activate();
    await context.sync();
    return { ok: true, action: "create_sheet", name };
  });
}

async function deleteSheet(name: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, name);
    sheet.delete();
    await context.sync();
    return { ok: true, action: "delete_sheet", name };
  });
}

async function activateSheet(name: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, name);
    sheet.activate();
    await context.sync();
    return { ok: true, action: "activate_sheet", name };
  });
}

async function insertTable(sheetName: string, address: string, hasHeaders: boolean): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const table = sheet.tables.add(address, hasHeaders);
    const tableRange = table.getRange();
    const strictA1Dimensions = parseA1RangeDimensions(address);
    const responseAddress = strictA1Dimensions
      ? `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`
      : undefined;

    table.style = "TableStyleMedium2";
    table.load("name");
    if (responseAddress === undefined) {
      tableRange.load("address");
    }
    tableRange.select();
    await context.sync();
    return {
      ok: true,
      action: "insert_table",
      tableName: table.name,
      address: responseAddress ?? tableRange.address
    };
  });
}

async function sortTable(tableName: string, columnIndex: number, ascending: boolean): Promise<string> {
  return safelyExcel(async (context) => {
    const table = context.workbook.tables.getItem(tableName);
    table.sort.apply([{ key: columnIndex, ascending }]);
    table.getRange().select();
    await context.sync();
    return { ok: true, action: "sort_table", tableName, columnIndex, ascending };
  });
}

async function filterTable(tableName: string, columnIndex: number, args: Record<string, unknown>): Promise<string> {
  return safelyExcel(async (context) => {
    const table = context.workbook.tables.getItem(tableName);
    const filter = table.columns.getItemAt(columnIndex).filter;
    if (Array.isArray(args.values)) {
      filter.applyValuesFilter(args.values.map(String));
    } else if (typeof args.criterion1 === "string") {
      filter.applyCustomFilter(args.criterion1, typeof args.criterion2 === "string" ? args.criterion2 : undefined, (args.operator || "And") as Excel.FilterOperator);
    } else {
      filter.clear();
    }
    table.getRange().select();
    await context.sync();
    return { ok: true, action: "filter_table", tableName, columnIndex };
  });
}

async function createChart(sheetName: string, dataRange: string, chartType: string, title: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const source = sheet.getRange(dataRange);

    const chart = sheet.charts.add(chartType as Excel.ChartType, source, "Auto");
    chart.title.text = title;
    chart.legend.position = "Right";

    // Derive the source range from the user input to avoid unnecessary readback.
    const normalizedSource = stripSheetName(dataRange).replace(/\$/g, "").trim();

    // Cache getLastColumn() — both offset calls reference the same column proxy.
    const lastCol = source.getLastColumn();
    const topLeft = lastCol.getOffsetRange(0, 2);
    const bottomRight = lastCol.getOffsetRange(14, 9);
    chart.setPosition(topLeft, bottomRight);
    chart.activate();
    await context.sync();

    return { ok: true, action: "create_chart", sheet: sheetName, source: normalizedSource, chartType, title };
  });
}

async function findReplace(sheetName: string, find: string, replace: string, matchCase: boolean): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getUsedRangeOrNullObject();
    // Load the null-object flag before reading it — Office.js proxies require sync
    range.load("isNullObject");
    await context.sync();
    if (range.isNullObject) {
      return {
        ok: true,
        action: "find_replace",
        changed: 0,
        message: "Sheet has no used range."
      };
    }

    range.load("address");
    const result = range.replaceAll(find, replace, { matchCase });
    await context.sync();

    return {
      ok: true,
      action: "find_replace",
      sheet: sheetName,
      address: stripSheetName(range.address),
      changed: result.value ?? 0
    };
  });
}

async function clearRange(sheetName: string, address: string): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    range.clear("All");
    range.select();
    const quotedSheetName = formatSheetNameForFormula(sheetName);
    const rangeAddress = stripSheetName(address).replace(/\$/g, "").trim();
    await context.sync();
    return {
      ok: true,
      action: "clear_range",
      sheet: sheetName,
      address: `${quotedSheetName}!${rangeAddress}`
    };
  });
}

async function runCalculation(): Promise<string> {
  return safelyExcel(async (context) => {
    context.workbook.application.calculate("Full");
    await context.sync();
    return { ok: true, action: "run_calculation" };
  });
}

async function conditionalFormatRange(sheetName: string, address: string, ruleType: string, args: Record<string, unknown>): Promise<string> {
  return safelyExcel(async (context) => {
    const range = getSheet(context, sheetName).getRange(address);
    const deterministicAddress = parseA1RangeDimensions(address)
      ? `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`
      : null;
    if (deterministicAddress === null) {
      range.load("address");
    }

    const normalizedRuleType = parseConditionalFormatRuleType(ruleType);
    const fillColor = optionalTextArg(args, "fillColor");
    const fontColor = optionalTextArg(args, "fontColor");

    let conditionalRule: any;
    if (normalizedRuleType === "cellValue") {
      const operator = optionalTextArg(args, "operator") ?? "GreaterThan";
      const formula1 = textArg(args, "formula1");
      const formula2 = optionalTextArg(args, "formula2");
      const conditionalFormat = range.conditionalFormats.add("CellValue" as Excel.ConditionalFormatType) as unknown as Record<string, any>;
      conditionalRule = conditionalFormat.cellValue;
      if (!conditionalRule) {
        throw new Error("Cell-value conditional formatting is not available on this device.");
      }
      conditionalRule.rule = { operator, formula1, formula2 };
    } else if (normalizedRuleType === "containsText") {
      const operator = optionalTextArg(args, "operator") ?? "Contains";
      const text = textArg(args, "text");
      const conditionalFormat = range.conditionalFormats.add("ContainsText" as Excel.ConditionalFormatType) as unknown as Record<string, any>;
      conditionalRule = conditionalFormat.textComparison;
      if (!conditionalRule) {
        throw new Error("Text conditional formatting is not available on this device.");
      }
      conditionalRule.rule = { operator, text };
    } else {
      throw new Error(`Conditional format rule type "${normalizedRuleType}" is not implemented yet in this build.`);
    }

    if (fillColor !== undefined) {
      conditionalRule.format.fill.color = fillColor;
    }
    if (fontColor !== undefined) {
      conditionalRule.format.font.color = fontColor;
    }
    await context.sync();
    return { ok: true, action: "conditional_format_range", sheet: sheetName, address: deterministicAddress ?? range.address, ruleType: normalizedRuleType };
  });
}

async function freezePanes(sheetName: string, rows: number, columns: number): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    sheet.activate();

    if (rows === 0 && columns === 0) {
      sheet.freezePanes.unfreeze();
      await context.sync();
      return { ok: true, action: "freeze_panes", sheet: sheetName, rows, columns, status: "unfrozen" };
    }

    if (rows > 0) {
      sheet.freezePanes.freezeRows(rows);
    }
    if (columns > 0) {
      sheet.freezePanes.freezeColumns(columns);
    }
    await context.sync();
    return { ok: true, action: "freeze_panes", sheet: sheetName, rows, columns, status: "frozen" };
  });
}

async function setDataValidation(sheetName: string, address: string, type: string, source: string, promptTitle?: string, promptMessage?: string): Promise<string> {
  return safelyExcel(async (context) => {
    const validationType = parseDataValidationType(type);
    const validationRange = getSheet(context, sheetName).getRange(address);
    const deterministicAddress = parseA1RangeDimensions(address)
      ? `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`
      : null;

    if (validationType !== "list") {
      throw new Error(`Unsupported data validation type "${validationType}". Supported: list.`);
    }

    const listSource = parseDataValidationSource(source);
    const validation = validationRange.dataValidation;
    (validation as unknown as Excel.DataValidation).ignoreBlanks = true;
    validation.rule = {
      list: {
        inCellDropDown: true,
        source: listSource
      }
    };

    if (promptTitle !== undefined || promptMessage !== undefined) {
      validation.prompt = {
        showPrompt: true,
        title: promptTitle ?? "",
        message: promptMessage ?? ""
      };
    }
    validationRange.select();
    if (deterministicAddress === null) {
      validationRange.load("address");
    }
    await context.sync();

    return { ok: true, action: "set_data_validation", sheet: sheetName, address: deterministicAddress ?? validationRange.address, type: validationType };
  });
}

async function autofillFormula(sheetName: string, sourceAddress: string, targetAddress: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const source = sheet.getRange(sourceAddress);
    const target = sheet.getRange(targetAddress);
    const sourceDimensions = parseA1RangeDimensions(sourceAddress);
    const targetDimensions = parseA1RangeDimensions(targetAddress);
    const response = { ok: true, action: "autofill_formula", sheet: sheetName, source: sourceAddress, target: targetAddress };

    // Strict A1 inputs provide enough local metadata for the normal autoFill path.
    // Queue the autofill mutation directly so it completes in one host round trip.
    if (
      sourceDimensions?.rows === 1 &&
      sourceDimensions.columns === 1 &&
      targetDimensions !== null
    ) {
      try {
        source.autoFill(target);
        target.select();
        await context.sync();
      } catch {
        // Fall through to the formula-copy fallback after loading only the needed formula.
        source.load("formulasR1C1");
        await context.sync();

        const r1c1 = (source.formulasR1C1 as unknown[][])[0][0];
        if (typeof r1c1 !== "string" || !r1c1.startsWith("=")) {
          throw new Error("Source cell does not contain a formula.");
        }
        target.formulasR1C1 = Array.from({ length: targetDimensions.rows }, () =>
          Array.from({ length: targetDimensions.columns }, () => r1c1)
        );
        target.select();
        await context.sync();
      }
      return response;
    }

    // Non-A1 inputs retain metadata validation and the existing fallback behavior.
    source.load(["formulasR1C1", "formulas", "rowCount", "columnCount"]);
    target.load(["address", "rowCount", "columnCount"]);
    await context.sync();

    if (source.rowCount !== 1 || source.columnCount !== 1) {
      throw new Error("Source address must be a single cell.");
    }

    try {
      source.autoFill(target);
      target.select();
      await context.sync();
    } catch {
      // Fallback: use formulasR1C1 so relative references adjust per target cell.
      // R1C1 formulas (e.g. "=R[-1]C*2") are interpreted relative to *each* cell's position,
      // which is the correct semantics for fill operations when autoFill is unavailable.
      const r1c1 = (source.formulasR1C1 as unknown[][])[0][0];
      if (typeof r1c1 !== "string" || !r1c1.startsWith("=")) {
        throw new Error("Source cell does not contain a formula.");
      }
      target.formulasR1C1 = Array.from({ length: target.rowCount }, () =>
        Array.from({ length: target.columnCount }, () => r1c1)
      );
      target.select();
      await context.sync();
    }

    return response;
  });
}

async function renameSheet(oldName: string, newName: string): Promise<string> {
  if (oldName === newName) {
    return JSON.stringify({ ok: true, action: "rename_sheet", oldName, newName });
  }

  return safelyExcel(async (context) => {
    const sheet = getSheet(context, oldName);
    sheet.name = newName;
    await context.sync();
    return { ok: true, action: "rename_sheet", oldName, newName };
  });
}

async function copySheet(sourceName: string, newName: string): Promise<string> {
  if (sourceName === newName) {
    return JSON.stringify({
      ok: false,
      error: "Source and destination sheet names must be different."
    });
  }

  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sourceName);
    const copied = sheet.copy("After", sheet);
    copied.name = newName;
    copied.activate();
    await context.sync();
    return { ok: true, action: "copy_sheet", sourceName, newName };
  });
}

/** Protection options known to work on Excel desktop even if absent from the type def. */
interface SheetProtectionOptions {
  allowSort?: boolean;
  allowAutoFilter?: boolean;
}

async function protectSheet(sheetName: string, password?: string, allowSort?: boolean, allowFilter?: boolean): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const options: SheetProtectionOptions = {};
    if (allowSort !== undefined) options.allowSort = allowSort;
    if (allowFilter !== undefined) options.allowAutoFilter = allowFilter;
    sheet.protection.protect(options as unknown as Excel.WorksheetProtectionOptions, password || "");
    await context.sync();
    return { ok: true, action: "protect_sheet", sheet: sheetName };
  });
}

async function unprotectSheet(sheetName: string, password?: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    sheet.protection.unprotect(password || "");
    await context.sync();
    return { ok: true, action: "unprotect_sheet", sheet: sheetName };
  });
}

async function createPivotTable(
  targetSheet: string, targetAddress: string,
  sourceSheet: string, sourceRange: string,
  rowFields?: string[], columnFields?: string[],
  dataFields?: Array<Record<string, string>>
): Promise<string> {
  return safelyExcel(async (context) => {
    const target = getSheet(context, targetSheet);
    const targetCell = target.getRange(targetAddress);

    // The tool inputs already provide the source worksheet/range and destination address.
    // Avoid loading proxy metadata before queuing the pivot mutation.

    // Check pivot API availability — fail fast with a clear message
    try {
      const _check = target.pivotTables; // eslint-disable-line @typescript-eslint/no-unused-vars
    } catch {
      throw new Error("Pivot tables are not supported in this Excel version through Office.js. Try summarizing data with summarize_range_stats or building a manual SUMIFS table.");
    }

    // The generated name is already known locally, so hierarchy configuration can
    // remain in the same batch as pivot creation.
    const pivotTableName = `PivotTable${Date.now()}`;
    // Excel requires worksheet names with special characters to be single-quoted;
    // embedded apostrophes are escaped by doubling them. Preserve the compact
    // reference form for simple names to keep the existing response contract.
    const pivotSourceSheet = formatSheetNameForFormula(sourceSheet);
    const pivotSource = `${pivotSourceSheet}!${sourceRange}`;
    const pivotTable = target.pivotTables.add(
      pivotTableName,
      pivotSource,
      targetCell
    );

    const appliedRowFields: string[] = [];
    const appliedColumnFields: string[] = [];
    const appliedDataFields: Array<{ name: string; summarizeBy: string }> = [];

    // Configure row fields — batch all lookups then single sync
    if (rowFields) {
      for (const fieldName of rowFields) {
        try {
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
          appliedRowFields.push(fieldName);
        } catch (e) {
          throw new Error(`Row field "${fieldName}" not found in source data or cannot be added. Available hierarchies may differ from column headers.`);
        }
      }
    }

    // Configure column fields
    if (columnFields) {
      for (const fieldName of columnFields) {
        try {
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
          appliedColumnFields.push(fieldName);
        } catch (e) {
          throw new Error(`Column field "${fieldName}" not found in source data or cannot be added.`);
        }
      }
    }

    // Configure data fields
    if (dataFields) {
      for (const df of dataFields) {
        const name = df.name;
        const summarizeBy = (df.summarizeBy || "sum").toLowerCase();
        try {
          const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(name));
          dataHierarchy.summarizeBy = summarizeBy as any;
          appliedDataFields.push({ name, summarizeBy });
        } catch (e) {
          throw new Error(`Data field "${name}" not found in source data or cannot be added.`);
        }
      }
    }

    // Finalize hierarchy mutations and target activation in one batch.
    target.activate();
    await context.sync();

    return {
      ok: true,
      action: "create_pivot_table",
      pivotTableName,
      source: pivotSource,
      targetSheet,
      targetCell: targetAddress,
      rowFields: appliedRowFields,
      columnFields: appliedColumnFields,
      dataFields: appliedDataFields
    };
  });
}

/**
 * Add or update a comment (note) on a specific cell.
 * Detects whether the Excel host supports the Comments API (ExcelApi 1.10+)
 * and falls back to a clear error message on unsupported hosts.
 */
async function addComment(sheetName: string, address: string, text: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const range = sheet.getRange(address);
    const canonicalAddress = `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`;

    // Check Comments API availability — ExcelApi 1.10+ required
    try {
      const _check = sheet.comments; // eslint-disable-line @typescript-eslint/no-unused-vars
    } catch {
      throw new Error(
        "Comments are not supported in this Excel host through Office.js. " +
        "The Comments API (ExcelApi 1.10+) requires a recent version of Excel desktop or Excel on the web."
      );
    }

    // getItemByCell returns a proxy; ItemNotFound surfaces async during sync.
    // Load the comment first so we can distinguish existing vs missing.
    let updated = false;
    const existing = sheet.comments.getItemByCell(range);
    existing.load("content");
    try {
      await context.sync();
      // Comment exists — queue the content update
      existing.content = text;
      updated = true;
    } catch (error) {
      // ItemNotFound → no existing comment — create a new one
      if (isItemNotFoundError(error)) {
        sheet.comments.add(range, text);
      } else {
        throw error;
      }
    }

    // Flush update or create mutation
    await context.sync();

    return { ok: true, action: "add_comment", sheet: sheetName, address: canonicalAddress, updated };
  });
}

/**
 * Delete all comments (notes) in a range.
 * Detects host support and returns a clear error on unsupported hosts.
 */
async function clearComments(sheetName: string, address: string): Promise<string> {
  return safelyExcel(async (context) => {
    const sheet = getSheet(context, sheetName);
    const range = sheet.getRange(address);
    const deterministicAddress = parseA1RangeDimensions(address)
      ? `${formatSheetNameForFormula(sheetName)}!${stripSheetName(address).replace(/\$/g, "").trim()}`
      : null;

    // Check Comments API availability before queueing work.
    // ExcelApi 1.10+ required.
    try {
      const _check = sheet.comments; // eslint-disable-line @typescript-eslint/no-unused-vars
    } catch {
      throw new Error(
        "Comments are not supported in this Excel host through Office.js. " +
        "The Comments API (ExcelApi 1.10+) requires a recent version of Excel desktop or Excel on the web."
      );
    }

    const comments = sheet.comments;

    // Queue comment metadata and only load the range address when the input is
    // not a strict A1 range. Strict A1 inputs already determine the response
    // address locally, avoiding unnecessary host hydration.
    comments.load("items");
    if (deterministicAddress === null) {
      range.load("address");
    }
    await context.sync();

    const commentItems = comments.items ?? [];
    if (commentItems.length === 0) {
      return {
        ok: true,
        action: "clear_comments",
        sheet: sheetName,
        address: deterministicAddress ?? range.address,
        deletedCount: 0
      };
    }

    // Load the address of each comment's cell via the typed getLocation() API.
    // getLocation() returns a Range; we load "address" on each one.
    const locations: Excel.Range[] = [];
    for (const comment of commentItems) {
      const loc = comment.getLocation();
      loc.load("address");
      locations.push(loc);
    }
    await context.sync();

    // Now addresses are populated — filter and delete
    const rangeAddress = deterministicAddress ?? range.address;
    const toDelete = commentItems.filter((_: unknown, i: number) =>
      isAddressInRange(locations[i].address, rangeAddress)
    );

    if (toDelete.length === 0) {
      return {
        ok: true,
        action: "clear_comments",
        sheet: sheetName,
        address: deterministicAddress ?? range.address,
        deletedCount: 0
      };
    }

    for (const comment of toDelete) {
      comment.delete();
    }
    await context.sync();

    return { ok: true, action: "clear_comments", sheet: sheetName, address: deterministicAddress ?? range.address, deletedCount: toDelete.length };
  });
}

/**
 * Pure: normalize CSV rows to equal column count (pad with empty strings).
 * Returns a 2D string array. Does not depend on Excel context.
 */
export function normalizeCsvRows(data: unknown[][], totalCols: number): string[][] {
  return data.map((row: unknown[]) => {
    const strRow = row.map((cell: unknown) =>
      cell === null || cell === undefined ? "" : String(cell)
    );
    while (strRow.length < totalCols) {
      strRow.push("");
    }
    return strRow;
  });
}

/**
 * Parse a CSV string (handling quoted commas, escaped quotes, CRLF, blank fields,
 * and malformed CSV) and write the resulting data to a worksheet starting at the
 * given address.
 */
async function importCsvToSheet(sheetName: string, address: string, csv: string, hasHeaders: boolean): Promise<string> {
  return safelyExcel(async (context) => {
    // Parse the CSV — PapaParse handles RFC 4180 edge cases robustly
    const result = Papa.parse(csv, {
      header: false,
      skipEmptyLines: false,
      dynamicTyping: false,  // keep everything as strings for Office.js
      delimitersToGuess: [",", "\t", ";", "|"]
    });

    if (result.errors && result.errors.length > 0) {
      // Malformed CSV: fail fast with a clear message instead of writing partial/corrupt data.
      // Exclude "Delimiter" errors — they happen for empty input where no delimiter
      // can be detected; the empty-data check below handles that case with a clear message.
      const hasStructuralError = result.errors.some(
        (e: Papa.ParseError) => e.type === "Quotes" || e.type === "FieldMismatch"
      );
      if (!result.data || (result.data as unknown[][]).length === 0 || hasStructuralError) {
        const errorLines = result.errors.map((e: Papa.ParseError) =>
          `Row ${e.row !== undefined ? e.row + 1 : "?"}: ${e.message}`
        );
        throw new Error(`CSV parsing failed: ${errorLines.join("; ")}`);
      }
    }

    const data = result.data as unknown[][];
    if (data.length === 0) {
      throw new Error("CSV data is empty. No rows to import.");
    }

    // Trim only completely blank rows at the end (e.g. trailing newline artifacts from PapaParse).
    // Blank rows inside the CSV are meaningful row positions and must be preserved.
    // Must compute column count from the cleaned data — discarded trailing blank rows
    // can be wider than the retained data and would otherwise inflate the write range.
    const cleanedData = data.slice();
    while (cleanedData.length > 0) {
      const lastRow = cleanedData[cleanedData.length - 1];
      const hasContent = lastRow.some((cell) => cell !== "" && cell !== null && cell !== undefined);
      if (hasContent) break;
      cleanedData.pop();
    }
    if (cleanedData.length === 0) {
      throw new Error("CSV data contains only blank rows. Nothing to import.");
    }

    const totalCols = Math.max(...cleanedData.map((row) => row.length), 0);

    if (totalCols === 0) {
      throw new Error("CSV data has no columns to import.");
    }

    // Use cleaned row count for range sizing so getResizedRange matches
    // the normalized data that will actually be written.
    const cleanedRows = cleanedData.length;

    // Normalize rows to equal column count (pad with empty strings)
    const normalizedData = normalizeCsvRows(cleanedData, totalCols);

    const sheet = getSheet(context, sheetName);
    const target = sheet.getRange(address);

    // CSV dimensions are known locally; build the final write range in the mutation batch.
    // Strict A1 inputs also provide a deterministic response address, avoiding host
    // address hydration. Non-A1 expressions retain the metadata fallback.
    const writeRange = target.getResizedRange(cleanedRows - 1, totalCols - 1);
    const deterministicAddress = getResizedA1Address(sheetName, address, cleanedRows, totalCols);
    if (deterministicAddress === null) {
      writeRange.load("address");
    }
    writeRange.values = normalizedData;

    // Apply formatting — bold headers row if hasHeaders
    if (hasHeaders && cleanedRows > 0) {
      const headerRange = writeRange.getCell(0, 0).getResizedRange(0, totalCols - 1);
      headerRange.format.font.bold = true;
    }

    await context.sync();

    return {
      ok: true,
      action: "import_csv_to_sheet",
      sheet: sheetName,
      address: deterministicAddress ?? writeRange.address,
      rows: cleanedRows,
      columns: totalCols,
      hasHeaders,
      parseErrors: result.errors && result.errors.length > 0
        ? result.errors.map((e: Papa.ParseError) => `Row ${e.row !== undefined ? e.row + 1 : "?"}: ${e.message}`)
        : []
    };
  });
}

/**
 * Check whether a cell address (e.g. "Sheet1!C5" or "C5") falls within
 * a target A1-style range address (e.g. "C2:D10").
 * Handles empty/malformed inputs gracefully (returns false).
 * Strips "$" from absolute reference notation (e.g. "$A$1" -> "A1").
 */
export function isAddressInRange(cellAddress: string, rangeAddress: string): boolean {
  if (!cellAddress || !rangeAddress) return false;

  // Strip sheet name prefix if present (handle cell and range independently)
  const cellPart = cellAddress.includes("!")
    ? cellAddress.split("!").slice(1).join("!")
    : cellAddress;
  const cleanedCell = cellPart.replace(/\$/g, "");
  const rangePart = rangeAddress.includes("!")
    ? rangeAddress.split("!").slice(1).join("!")
    : rangeAddress;
  const cleanedRange = rangePart.replace(/\$/g, "");

  const cellLetter = cleanedCell.replace(/[0-9]/g, "");
  const cellNumber = cleanedCell.replace(/[A-Z]/gi, "");
  if (!cellLetter || !cellNumber) return false;

  const cellCol = parseColumnLabel(cellLetter);
  const cellRow = parseInt(cellNumber, 10);
  if (isNaN(cellRow)) return false;

  const [startAddr, endAddr] = cleanedRange.split(":").map((s) => s.trim());
  const startLetter = startAddr.replace(/[0-9]/g, "");
  const startNumber = startAddr.replace(/[A-Z]/gi, "");
  if (!startLetter || !startNumber) return false;

  const startCol = parseColumnLabel(startLetter);
  const startRow = parseInt(startNumber, 10);
  if (isNaN(startRow)) return false;

  const endLetter = endAddr ? endAddr.replace(/[0-9]/g, "") : startLetter;
  const endNumber = endAddr ? endAddr.replace(/[A-Z]/gi, "") : startNumber;
  if (!endLetter || !endNumber) return false;

  const endCol = parseColumnLabel(endLetter);
  const endRow = parseInt(endNumber, 10);
  if (isNaN(endRow)) return false;

  return cellCol >= startCol && cellCol <= endCol && cellRow >= startRow && cellRow <= endRow;
}

/**
 * Convert a column label (e.g. "A", "Z", "AA") to a 1-based column number.
 */
export function parseColumnLabel(label: string): number {
  // Normalize to uppercase so that callers can pass lowercase
  // addresses (e.g. clearComments with deterministicAddress from
  // a user-supplied lowercase A1 range) without breaking the
  // charCode-based arithmetic below.
  const upper = label.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);  // "A" = 1
  }
  return result;
}

function getPreviewA1Address(sheetName: string, address: string, rows: number, columns: number): string | null {
  const rangePart = stripSheetName(address).replace(/\$/g, "").trim();
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(rangePart);
  if (!match || rows < 1 || columns < 1) {
    return null;
  }

  const startColumn = parseColumnLabel(match[1].toUpperCase());
  const startRow = Number(match[2]);
  const endColumn = startColumn + columns - 1;
  const endRow = startRow + rows - 1;
  if (
    !Number.isSafeInteger(startColumn) ||
    !Number.isSafeInteger(startRow) ||
    startRow < 1 ||
    !Number.isSafeInteger(endColumn) ||
    !Number.isSafeInteger(endRow) ||
    endRow < startRow ||
    endColumn < startColumn
  ) {
    return null;
  }

  const startAddress = `${columnNumberToLabel(startColumn)}${startRow}`;
  const endAddress = `${columnNumberToLabel(endColumn)}${endRow}`;
  const previewRange = startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
  return `${formatSheetNameForFormula(sheetName)}!${previewRange}`;
}

function getResizedA1Address(sheetName: string, address: string, rows: number, columns: number): string | null {
  const rangePart = stripSheetName(address).replace(/\$/g, "").trim();
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(rangePart);
  if (!match || rows < 1 || columns < 1) {
    return null;
  }

  const startColumn = parseColumnLabel(match[1].toUpperCase());
  const startRow = Number(match[2]);
  const targetEndColumn = match[3] ? parseColumnLabel(match[3].toUpperCase()) : startColumn;
  const targetEndRow = match[4] ? Number(match[4]) : startRow;
  const targetRows = targetEndRow - startRow + 1;
  const targetColumns = targetEndColumn - startColumn + 1;
  if (
    !Number.isSafeInteger(startRow) ||
    startRow < 1 ||
    !Number.isSafeInteger(targetEndRow) ||
    targetRows < 1 ||
    targetColumns < 1
  ) {
    return null;
  }

  const endColumn = startColumn + targetColumns + columns - 2;
  const endRow = startRow + targetRows + rows - 2;
  if (!Number.isSafeInteger(endRow) || endRow < startRow || endColumn < startColumn) {
    return null;
  }

  const startAddress = `${columnNumberToLabel(startColumn)}${startRow}`;
  const endAddress = `${columnNumberToLabel(endColumn)}${endRow}`;
  const finalRange = startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
  return `${formatSheetNameForFormula(sheetName)}!${finalRange}`;
}

function columnNumberToLabel(columnNumber: number): string {
  let remaining = columnNumber;
  let label = "";
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return label;
}

function parseA1RangeDimensions(address: string): { rows: number; columns: number } | null {
  const rangePart = stripSheetName(address).replace(/\$/g, "").trim();
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(rangePart);
  if (!match) {
    return null;
  }

  const startColumn = parseColumnLabel(match[1].toUpperCase());
  const startRow = Number(match[2]);
  const endColumn = match[3] ? parseColumnLabel(match[3].toUpperCase()) : startColumn;
  const endRow = match[4] ? Number(match[4]) : startRow;
  if (!Number.isSafeInteger(startRow) || !Number.isSafeInteger(endRow) || startRow < 1 || endRow < startRow || endColumn < startColumn) {
    return null;
  }

  return {
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1
  };
}

function getSheet(context: Excel.RequestContext, sheetName: string): Excel.Worksheet {
  return context.workbook.worksheets.getItem(sheetName);
}

function errorDetails(error: unknown): { code: string; message: string; name: string } {
  if (typeof error === "object" && error !== null) {
    const record = error as Record<string, unknown>;
    return {
      code: typeof record.code === "string" ? record.code : "",
      message: typeof record.message === "string" ? record.message : error instanceof Error ? error.message : String(error),
      name: typeof record.name === "string" ? record.name : error instanceof Error ? error.name : "Error"
    };
  }
  return { code: "", message: String(error), name: "Error" };
}

function isItemNotFoundError(error: unknown): boolean {
  const { code, message } = errorDetails(error);
  return code === "ItemNotFound" || message.includes("ItemNotFound") || message.includes("item not found");
}

/** R4e-2: monotonic count of Office batches whose work callback started (read-only; any process-wide batch counts). */
let officeBatchEntries = 0;
export function getOfficeBatchEntryCount(): number {
  return officeBatchEntries;
}

async function safelyExcel(work: (context: Excel.RequestContext) => Promise<Record<string, unknown>>): Promise<string> {
  try {
    const result = await Excel.run((context) => { officeBatchEntries += 1; return work(context); });
    return JSON.stringify(result);
  } catch (error) {
    // Translate common Office.js errors into user-friendly messages
    const details = errorDetails(error);
    if (isItemNotFoundError(error)) {
      return JSON.stringify({ ok: false, error: `Sheet not found. Verify the sheet name is correct and try again.`, name: details.name });
    }
    if (details.message.includes("already exists") || details.message.includes("name conflict") || details.message.includes("Conflict")) {
      return JSON.stringify({ ok: false, error: `A sheet or item with that name already exists. Please choose a different name.`, name: details.name });
    }
    return stringifyError(error);
  }
}

function sheetTool(name: string, description: string, required: string[], execute: (args: Record<string, unknown>) => Promise<string>): ExcelToolDefinition {
  return {
    schema: {
      type: "function",
      function: {
        name,
        description,
        parameters: {
          type: "object",
          properties: {
            name: { type: "string" }
          },
          required,
          additionalProperties: false
        }
      }
    },
    execute
  };
}

function textArg(args: Record<string, unknown>, name: string): string {
  const value = args[name];
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`Missing or invalid string argument "${name}".`);
  }
  return value;
}

function boolArg(args: Record<string, unknown>, name: string): boolean {
  const value = args[name];
  if (typeof value !== "boolean") {
    throw new Error(`Missing or invalid boolean argument "${name}".`);
  }
  return value;
}

export function parseConditionalFormatRuleType(raw: unknown): ConditionalFormatRuleType {
  if (typeof raw !== "string") {
    throw new Error(`Missing or invalid string argument "ruleType".`);
  }

  const normalized = raw.trim();
  const allowed: ConditionalFormatRuleType[] = ["cellValue", "containsText"];
  if (!allowed.includes(normalized as ConditionalFormatRuleType)) {
    throw new Error(`Unsupported conditional format ruleType "${raw}".`);
  }

  return normalized as ConditionalFormatRuleType;
}

export function parseDataValidationType(raw: unknown): DataValidationType {
  if (typeof raw !== "string") {
    throw new Error(`Missing or invalid string argument "type".`);
  }

  const normalized = raw.trim();
  if (normalized !== "list") {
    throw new Error(`Unsupported data validation type "${raw}".`);
  }

  return normalized as DataValidationType;
}

export function parseDataValidationSource(raw: unknown): string {
  if (typeof raw !== "string") {
    throw new Error(`Missing or invalid string argument "source".`);
  }
  const normalized = raw.trim();
  if (normalized === "") {
    throw new Error("Missing or invalid string argument \"source\".");
  }
  return normalized;
}

export function nonNegativeIntArg(args: Record<string, unknown>, name: string): number {
  const value = args[name];
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
    throw new Error(`Missing or invalid non-negative integer argument "${name}".`);
  }
  return value;
}

function numberArg(args: Record<string, unknown>, name: string): number {
  const value = Number(args[name]);
  if (!Number.isFinite(value)) {
    throw new Error(`Missing or invalid number argument "${name}".`);
  }
  return Math.floor(value);
}

function objectArg(args: Record<string, unknown>, name: string): Record<string, unknown> {
  const value = args[name];
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`Missing or invalid object argument "${name}".`);
  }
  return value as Record<string, unknown>;
}

function optionalTextArg(args: Record<string, unknown>, name: string): string | undefined {
  const value = args[name];
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new Error(`Invalid string argument "${name}".`);
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function optionalBoolArg(args: Record<string, unknown>, name: string): boolean | undefined {
  if (args[name] === undefined || args[name] === null) {
    return undefined;
  }
  if (typeof args[name] !== "boolean") {
    throw new Error(`Missing or invalid boolean argument "${name}".`);
  }
  return args[name] as boolean;
}

function array2dArg(args: Record<string, unknown>, name: string): JsonValue[][] {
  const value = args[name];
  if (!Array.isArray(value) || value.some((row) => !Array.isArray(row))) {
    throw new Error(`Missing or invalid 2D array argument "${name}".`);
  }
  return value as JsonValue[][];
}

export function parseNumericCell(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed === "") {
    return null;
  }

  const normalized = trimmed.replace(/,/g, "");
  if (!/^[-+]?(?:\d{1,3}(?:,\d{3})+|\d+|\.\d+)(?:\.\d+)?([eE][-+]?\d+)?$/.test(trimmed)) {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function isBlankValue(value: unknown): boolean {
  return value === null || value === undefined || (typeof value === "string" && value.trim() === "");
}

export function summarizeRangeValues(values: ReadonlyArray<ReadonlyArray<unknown>>): RangeStats {
  let count = 0;
  let numericCount = 0;
  let blankCount = 0;
  let sum = 0;
  let min: number | null = null;
  let max: number | null = null;

  for (const row of values) {
    for (const value of row) {
      count += 1;
      if (isBlankValue(value)) {
        blankCount += 1;
      }
      const parsed = parseNumericCell(value);
      if (parsed === null) {
        continue;
      }
      numericCount += 1;
      sum += parsed;
      min = min === null ? parsed : Math.min(min, parsed);
      max = max === null ? parsed : Math.max(max, parsed);
    }
  }

  const average = numericCount > 0 ? sum / numericCount : null;

  return {
    count,
    numericCount,
    sum,
    average,
    min,
    max,
    blankCount
  };
}

function assertDimensions(values: JsonValue[][], rows: number, columns: number): void {
  if (values.length !== rows || values.some((row) => row.length !== columns)) {
    throw new Error(`Range dimension mismatch. The target range is ${rows}x${columns}, but values are ${values.length}x${values[0]?.length ?? 0}.`);
  }
}

function stripSheetName(address: string): string {
  return address.includes("!") ? address.split("!").slice(1).join("!") : address;
}

function formatSheetNameForFormula(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replace(/'/g, "''")}'`;
}

function replaceText(value: string, find: string, replace: string, matchCase: boolean): string {
  if (!find) {
    return value;
  }
  if (matchCase) {
    return value.split(find).join(replace);
  }
  return value.replace(new RegExp(escapeRegExp(find), "gi"), replace);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function stringifyError(error: unknown): string {
  const details = errorDetails(error);
  return JSON.stringify({ ok: false, error: details.message, name: details.name });
}

// ---------------------------------------------------------------------------
// Selection address parser
// ---------------------------------------------------------------------------

export interface ParsedSelectionAddress {
  /** Unquoted/unescaped sheet name suitable for worksheets.getItem(). */
  sheet: string;
  /** Range address e.g. "A1:B10" or "A1". */
  range: string;
}

/**
 * Parse an Excel selection address into sheet and range components.
 *
 * Handles:
 *   - Unqualified ranges:              "A1:B10"
 *   - Ordinary sheet names:            "Sheet1!A1:B10"
 *   - Quoted sheet names with spaces:  "'My Sheet'!A1:B10"
 *   - Escaped apostrophes:            "'Jon''s Data'!A1"
 *   - Sheet names containing '!':     "'Data!2024'!A1:B10" (range is after last '!')
 *
 * The returned sheet is unquoted and unescaped. Callers must NOT
 * blindly split on '!' — use this function.
 */
export function parseExcelSelection(
  selection: string,
  fallbackSheet: string,
): ParsedSelectionAddress {
  const trimmed = selection.trim();
  if (!trimmed) {
    return { sheet: fallbackSheet, range: "" };
  }

  // If the address starts with a single quote, the sheet name is quoted
  // and may contain '!' — we need to find the closing quote.
  if (trimmed.startsWith("'")) {
    // Find the closing single quote. Escaped apostrophes inside the name
    // are doubled (''), so we look for an unescaped closing quote.
    let closeIdx = -1;
    for (let i = 1; i < trimmed.length; i++) {
      if (trimmed[i] === "'") {
        // Check if it's an escaped apostrophe (''  = two quotes in a row)
        if (i + 1 < trimmed.length && trimmed[i + 1] === "'") {
          i++; // skip the escaped quote
          continue;
        }
        closeIdx = i;
        break;
      }
    }

    if (closeIdx === -1) {
      // No closing quote found — treat whole thing as unquoted, no sheet
      return { sheet: fallbackSheet, range: trimmed };
    }

    // Sheet name is between the quotes (already unescaped)
    const rawSheet = trimmed.slice(1, closeIdx);
    const sheet = rawSheet.replace(/''/g, "'"); // unescape apostrophes

    // Everything after the closing quote should be "!" + range
    const afterQuote = trimmed.slice(closeIdx + 1);
    if (afterQuote.startsWith("!")) {
      return { sheet, range: afterQuote.slice(1) };
    }
    // Quoted name but no '!' separator — treat as unqualified range
    return { sheet, range: trimmed };
  }

  // Unquoted: find the LAST '!' — everything after it is the range.
  // Sheet names without special chars cannot contain '!', so the last
  // '!' always separates sheet from range.
  const lastBang = trimmed.lastIndexOf("!");
  if (lastBang === -1) {
    return { sheet: fallbackSheet, range: trimmed };
  }

  const sheet = trimmed.slice(0, lastBang);
  const range = trimmed.slice(lastBang + 1);
  return { sheet, range };
}
