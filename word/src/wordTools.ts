import { OpenAITool } from "./types";
import { categorizeError, ToolErrorCategory } from "./toolErrors";
import { parseAndValidateToolArguments } from "./toolValidationCore";

type JsonRecord = Record<string, unknown>;

type WordTool = {
  schema: OpenAITool;
  execute: (args: JsonRecord, context?: Word.RequestContext) => Promise<string>;
};

export interface ReadResult {
  text: string;
  truncated: boolean;
  originalLength: number;
  returnedLength: number;
}

export interface EditPlan {
  edits: EditPlanItem[];
}

export interface EditPlanItem {
  find: string;
  replace: string;
  matchCase: boolean;
  matchWholeWord: boolean;
  highlightColor?: string;
  reason?: string;
}

export interface ApplyEditPlanValidationSuccess {
  ok: true;
  plan: EditPlan;
}

export interface ApplyEditPlanValidationFailure {
  ok: false;
  error: string;
}

export interface ApplyEditPlanExecutionResult {
  ok: true;
  action: "apply_edit_plan";
  totalRequested: number;
  totalApplied: number;
  totalSkipped: number;
  edits: ApplyEditPlanExecutionItem[];
}

export interface ApplyEditPlanExecutionItem {
  index: number;
  status: "applied" | "not_found" | "ambiguous" | "overlapping";
  find: string;
  replace: string;
  matchCount: number;
  highlighted: boolean;
  reason: string | null;
}

export type ApplyEditPlanValidationResult = ApplyEditPlanValidationSuccess | ApplyEditPlanValidationFailure;

export interface DocumentSnapshot {
  overview: {
    charCount: number;
    wordCount: number;
    paragraphCount: number;
    selection: {
      start: number | null;
      end: number | null;
      text: string;
    };
  };
  documentText: ReadResult;
}

const MAX_EDIT_PLAN_EDITS = 100;
const MAX_FIND_LENGTH = 255;
const FIND_DESCRIPTION = "Literal source: at most 255 UTF-16 code units, including encoded query length (each caret encodes as ^^). Controls, tabs and paragraph/line breaks are unsupported. Never truncate or split.";
const PARAGRAPH_STYLES = ["Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6"];

type OperationState = { phase: string; mutationState: "none" | "unknown" };
class SearchValidationError extends Error {
  constructor(readonly detail: Record<string, unknown>) { super(String(detail.error)); }
}

function literalQuery(find: string, editIndex = 0): string {
  const encoded = find.replace(/\^/g, "^^");
  const error = find.length > MAX_FIND_LENGTH ? "search_source_too_long"
    : /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/.test(find) ? "unsupported_search_control"
    : encoded.length > MAX_FIND_LENGTH ? "search_encoded_too_long" : null;
  if (error) throw new SearchValidationError({ error, editIndex, sourceLength: find.length, encodedLength: encoded.length });
  return encoded;
}

function literalOptions(matchCase = false, matchWholeWord = false) {
  return { matchCase, matchWholeWord, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false };
}

function preflightSearch(name: string, args: JsonRecord): void {
  const sources = name === "apply_edit_plan" && Array.isArray(args.edits) ? args.edits.map(e => e?.find)
    : name === "find_replace" ? [args.find] : name === "apply_style" ? [args.targetText] : [];
  sources.forEach((source, index) => { if (typeof source === "string") literalQuery(source, index); });
}

function diagnosticFailure(error: unknown, state: OperationState): string {
  return JSON.stringify({ ok: false, error: "word_operation_failed", category: "office_host", ...state,
    ...(error instanceof SearchValidationError ? { ...error.detail, category: "validation", phase: "validation", mutationState: "none" } : {}) });
}
const MAX_REPLACE_LENGTH = 5000;
const MAX_OOXML_LENGTH = 20000;
const MAX_REASON_LENGTH = 250;

const MAX_READ_CHARS = 5000;
const CONTEXT_READ_CHARS = 4000;

export const toolAliases: Record<string, string> = {
  get_document_text: "read_document_text",
  read_document_range: "read_document_text"
};

export const wordTools: WordTool[] = [
  {
    schema: {
      type: "function",
      function: {
        name: "get_document_overview",
        description: "Get high-level details about the active Word document and current selection.",
        parameters: { type: "object", properties: {}, additionalProperties: false }
      }
    },
    execute: (_args, context) => getDocumentOverview(context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_selection",
        description: "Read the currently selected Word text.",
        parameters: { type: "object", properties: {}, additionalProperties: false }
      }
    },
    execute: (_args, context) => readSelection(context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_document_text",
        description: "Read text from the Word document body.",
        parameters: {
          type: "object",
          properties: {
            maxLength: { type: "integer", minimum: 1, maximum: 20000, description: "Maximum number of text characters to return." }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => readDocumentText(optionalBoundedPositiveIntegerArg(args, "maxLength", 20000) ?? MAX_READ_CHARS, context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "apply_edit_plan",
        description:
          "Validate and apply a bounded exact-match edit plan with optional highlighting. Edits must target independent, non-overlapping source text, and each source search must be unique for its match options. Returns one result per requested replacement.",
        parameters: {
          type: "object",
          properties: {
            edits: {
              type: "array",
              minItems: 1,
              maxItems: MAX_EDIT_PLAN_EDITS,
              items: {
                type: "object",
                properties: {
                  find: { type: "string", minLength: 1, maxLength: MAX_FIND_LENGTH, description: FIND_DESCRIPTION },
                  replace: { type: "string", maxLength: MAX_REPLACE_LENGTH },
                  reason: { type: "string", maxLength: MAX_REASON_LENGTH },
                  highlightColor: { type: "string", description: "Optional highlight color (yellow or hex)." },
                  matchCase: { type: "boolean", default: false },
                  matchWholeWord: { type: "boolean", default: false }
                },
                required: ["find", "replace"],
                additionalProperties: false
              }
            }
          },
          required: ["edits"],
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => applyEditPlan(args, context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "insert_text",
        description: "Insert text, or construct 1-20 native paragraphs with observed built-in styles (WordApi 1.3). Supply text OR paragraphs, never both. Paragraph mode accepts start/end only, total 5000 UTF-16 units; no embedded controls/breaks. Await and verify structure after insertion.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string", minLength: 1, maxLength: MAX_REPLACE_LENGTH },
            location: { type: "string", enum: ["start", "end", "selection"] },
            paragraphs: { type: "array", minItems: 1, maxItems: 20, items: { type: "object", properties: {
              text: { type: "string", minLength: 1, maxLength: MAX_REPLACE_LENGTH },
              styleBuiltIn: { type: "string", enum: PARAGRAPH_STYLES }
            }, required: ["text", "styleBuiltIn"], additionalProperties: false } }
          },
          minProperties: 1,
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => args.paragraphs !== undefined ? insertParagraphs(args, context)
      : insertText(nonEmptyTextArg(args, "text"), optionalTextArg(args, "location") ?? "end", context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "replace_selection",
        description: "Replace the current Word selection with text.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string", minLength: 1, maxLength: MAX_REPLACE_LENGTH },
            highlightColor: {
              type: "string",
              description: "Optional highlight color for replacement text, e.g. yellow or #FFFF00. Use yellow when the user asks to highlight changes."
            }
          },
          required: ["text"],
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => replaceSelection(nonEmptyTextArg(args, "text"), optionalTextArg(args, "highlightColor"), context)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "find_replace",
        description: "Find text in the Word document and replace matches.",
        parameters: {
          type: "object",
          properties: {
            find: { type: "string", minLength: 1, maxLength: MAX_FIND_LENGTH, description: FIND_DESCRIPTION },
            replace: { type: "string", maxLength: MAX_REPLACE_LENGTH },
            matchCase: { type: "boolean" },
            matchWholeWord: { type: "boolean", description: "Only replace whole-word matches." },
            highlightColor: {
              type: "string",
              description: "Optional highlight color for replacement text, e.g. yellow or #FFFF00. Use yellow when the user asks to highlight changes."
            }
          },
          required: ["find", "replace"],
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => findReplace(
      nonEmptyTextArg(args, "find"),
      textArg(args, "replace"),
      optionalBooleanArg(args, "matchCase") ?? false,
      optionalBooleanArg(args, "matchWholeWord") ?? false,
      optionalTextArg(args, "highlightColor"),
      context
    )
  },
  {
    schema: {
      type: "function",
      function: {
        name: "format_selection",
        description: "Set formatting properties on the current selection.",
        parameters: {
          type: "object",
          properties: {
            bold: { type: "boolean" },
            italic: { type: "boolean" },
            color: { type: "string", description: "Hex or named font color." },
            highlightColor: { type: "string", description: "Highlight color, e.g. yellow or #FFFF00." },
            fontSize: { type: "number", minimum: 1, description: "Font size in points." }
          },
          minProperties: 1,
          additionalProperties: false
        }
      }
    },
    execute: (args, context) => formatSelection(
      optionalBooleanArg(args, "bold"),
      optionalBooleanArg(args, "italic"),
      optionalTextArg(args, "color"),
      optionalTextArg(args, "highlightColor"),
      optionalMinimumNumberArg(args, "fontSize", 1),
      context
    )
  },
  {
    schema: { type: "function", function: {
      name: "get_document_structure",
      description: "Read paragraph counts, bounded whitespace-preserving previews and actual style/styleBuiltIn. styleBuiltIn is null with explicit capability status when unsupported; truncated previews are not exact target identities.",
      parameters: { type: "object", properties: { maxParagraphs: { type: "integer", minimum: 1, maximum: 100 } }, additionalProperties: false }
    } },
    execute: (args, context) => getDocumentStructure(optionalBoundedPositiveIntegerArg(args, "maxParagraphs", 100), context)
  },
  {
    schema: { type: "function", function: {
      name: "apply_style",
      description: "Apply style to the selection, or to one unique complete paragraph matching targetText exactly, independent of selection. Heading1..Heading6 and Normal use built-in styles (WordApi 1.3); other names use localized/custom style. Returns observed styles after sync. Missing, partial or ambiguous targets refuse without writes.",
      parameters: { type: "object", properties: { style: { type: "string", minLength: 1 }, targetText: { type: "string", minLength: 1, maxLength: MAX_FIND_LENGTH, description: FIND_DESCRIPTION } }, required: ["style"], additionalProperties: false }
    } },
    execute: (args, context) => applyStyle(nonEmptyTextArg(args, "style"), args.targetText === undefined ? undefined : nonEmptyTextArg(args, "targetText"), context)
  },
  {
    schema: { type: "function", function: {
      name: "insert_ooxml",
      description: "Insert trusted WordprocessingML at the document start, end, or current selection.",
      parameters: { type: "object", properties: { ooxml: { type: "string", minLength: 1, maxLength: MAX_OOXML_LENGTH }, location: { type: "string", enum: ["start", "end", "selection"] } }, required: ["ooxml"], additionalProperties: false }
    } },
    execute: (args, context) => insertOoxml(nonEmptyTextArg(args, "ooxml"), optionalTextArg(args, "location") ?? "end", context)
  },
  {
    schema: { type: "function", function: {
      name: "insert_comment",
      description: "Insert a comment on the current selection.",
      parameters: { type: "object", properties: { text: { type: "string", minLength: 1, maxLength: MAX_REPLACE_LENGTH } }, required: ["text"], additionalProperties: false }
    } },
    execute: (args, context) => insertComment(nonEmptyTextArg(args, "text"), context)
  }
];

export const toolSchemas = wordTools.map((tool) => tool.schema);

export function getSearchValidationFailure(name: string, rawArgs: string): string | undefined {
  const candidate = safeParseJson(rawArgs);
  try { if (candidate && typeof candidate === "object" && !Array.isArray(candidate)) preflightSearch(normalizeToolName(name), candidate as JsonRecord); }
  catch (error) { return diagnosticFailure(error, { phase: "validation", mutationState: "none" }); }
  return undefined;
}

export async function executeWordTool(name: string, rawArgs: string): Promise<string> {
  const canonicalName = normalizeToolName(name);
  const tool = wordTools.find((entry) => entry.schema.function.name === canonicalName);
  if (!tool) {
    return stringifyError(`Unknown tool: ${name}`, ToolErrorCategory.Validation);
  }

  const parameters = (tool.schema.function.parameters ?? { type: "object" }) as Record<string, unknown>;
  const searchFailure = getSearchValidationFailure(canonicalName, rawArgs);
  if (searchFailure) return searchFailure;
  const validation = parseAndValidateToolArguments(rawArgs, parameters);
  if (typeof validation === "string") {
    return stringifyError(new Error(formatDirectValidationError(canonicalName, validation)), ToolErrorCategory.Validation);
  }
  return executeWordToolParsed(canonicalName, validation.parsedArgs);
}

function formatDirectValidationError(toolName: string, validation: string): string {
  if (toolName === "format_selection" && validation === "arguments must contain at least 1 properties.") {
    return "format_selection requires at least one formatting property.";
  }
  const fieldMatch = validation.match(/^(?:[a-zA-Z_$][\w$]*(?:\[\d+\])?)(?:\.[a-zA-Z_$][\w$]*)*/);
  const fieldPath = fieldMatch?.[0];
  const field = fieldPath?.split(/[.\[]/)[0];
  if (!field) return `Tool "${toolName}" ${validation}`;
  if (validation.includes("must be at least 1 characters")) return `Argument "${field}" must not be empty.`;
  if (validation.includes("must not exceed ") && validation.includes("characters")) {
    return `Argument "${field}" ${validation.slice(validation.indexOf("must not exceed "))}`;
  }
  if (validation.includes("must be an integer") && (field === "maxLength" || field === "maxParagraphs")) {
    const maximum = field === "maxLength" ? 20000 : 100;
    return `Argument "${field}" must be an integer between 1 and ${maximum}.`;
  }
  if (validation.includes("must be at least ") && (field === "fontSize" || field === "maxLength" || field === "maxParagraphs")) {
    const maximum = field === "maxLength" ? 20000 : field === "maxParagraphs" ? 100 : undefined;
    return maximum === undefined
      ? `Argument "${field}" ${validation.slice(validation.indexOf("must be at least "))}`
      : `Argument "${field}" must be an integer between 1 and ${maximum}.`;
  }
  if (validation.includes("must be at most ") && (field === "maxLength" || field === "maxParagraphs")) {
    const maximum = field === "maxLength" ? 20000 : 100;
    return `Argument "${field}" must be an integer between 1 and ${maximum}.`;
  }
  if (validation.includes("must be one of") && field === "location") {
    return `${toolName} received an invalid location. Use "start", "end", or "selection".`;
  }
  return `Tool "${toolName}" ${validation}`;
}

/** Execute only an already parsed and schema-validated argument object. */
export async function executeWordToolParsed(name: string, parsedArgs: JsonRecord): Promise<string> {
  return executeParsedWithOptionalContext(name, parsedArgs);
}

/** Execute an already parsed tool in a caller-owned Word.run context. */
export async function executeWordToolInContext(
  name: string,
  parsedArgs: JsonRecord,
  context: Word.RequestContext
): Promise<string> {
  return executeParsedWithOptionalContext(name, parsedArgs, context);
}

async function executeParsedWithOptionalContext(
  name: string,
  parsedArgs: JsonRecord,
  context?: Word.RequestContext
): Promise<string> {
  const canonicalName = normalizeToolName(name);
  const tool = wordTools.find((entry) => entry.schema.function.name === canonicalName);
  if (!tool) {
    return stringifyError(`Unknown tool: ${name}`, ToolErrorCategory.Validation);
  }

  try {
    preflightSearch(canonicalName, parsedArgs);
    return addFailureCategory(await tool.execute(parsedArgs, context));
  } catch (error) {
    if (error instanceof SearchValidationError) return diagnosticFailure(error, { phase: "validation", mutationState: "none" });
    const category = categorizeError(error);
    return stringifyError(error, category === ToolErrorCategory.Internal && error instanceof Error && /argument|invalid|empty|required/i.test(error.message)
      ? ToolErrorCategory.Validation
      : category);
  }
}

function addFailureCategory(result: string): string {
  const parsed = safeParseJson(result);
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return result;
  const record = parsed as Record<string, unknown>;
  if (record.ok !== false || typeof record.category === "string") return result;
  const message = typeof record.error === "string" ? record.error : "Word tool failed.";
  const category = /requires WordApi|Word host|Office/i.test(message)
    ? ToolErrorCategory.OfficeHost
    : /edit|argument|invalid|duplicate|overlap|must include|requires/i.test(message)
      ? ToolErrorCategory.Validation
      : ToolErrorCategory.Internal;
  return JSON.stringify({ ...record, category });
}

function normalizeToolName(name: string): string {
  const normalized = name.trim().toLowerCase();
  return toolAliases[normalized] || normalized;
}

export function validateApplyEditPlan(rawArgs: unknown): ApplyEditPlanValidationResult {
  if (!rawArgs || typeof rawArgs !== "object" || Array.isArray(rawArgs)) {
    return { ok: false, error: "apply_edit_plan requires an object payload." };
  }

  const args = rawArgs as JsonRecord;
  const rawEdits = args.edits;
  if (!Array.isArray(rawEdits)) {
    return { ok: false, error: "apply_edit_plan requires an edits array." };
  }

  if (rawEdits.length < 1 || rawEdits.length > MAX_EDIT_PLAN_EDITS) {
    return {
      ok: false,
      error: `apply_edit_plan must include 1-${MAX_EDIT_PLAN_EDITS} edits.`
    };
  }

  const edits: EditPlanItem[] = [];
  const seenSearches = new Map<string, number>();

  for (let index = 0; index < rawEdits.length; index += 1) {
    const rawEdit = rawEdits[index];
    if (!rawEdit || typeof rawEdit !== "object" || Array.isArray(rawEdit)) {
      return { ok: false, error: `Edit ${index} is invalid.` };
    }

    const edit = rawEdit as Record<string, unknown>;
    const find = typeof edit.find === "string" ? edit.find : "";
    const rawReplace = edit.replace;
    if (typeof rawReplace !== "string") {
      return { ok: false, error: `Edit ${index} replace must be a string.` };
    }

    const replace = rawReplace;

    if (!find.trim()) {
      return { ok: false, error: `Edit ${index} is missing a non-empty find value.` };
    }
    try { literalQuery(find, index); }
    catch (error) { return { ok: false, error: error instanceof SearchValidationError ? String(error.detail.error) : "invalid_search" }; }

    if (replace === undefined) {
      return { ok: false, error: `Edit ${index} is missing a replace value.` };
    }
    if (typeof replace !== "string") {
      return { ok: false, error: `Edit ${index} replace must be a string.` };
    }
    if (replace.length > MAX_REPLACE_LENGTH) {
      return { ok: false, error: `Edit ${index} replacement text exceeds ${MAX_REPLACE_LENGTH} characters.` };
    }

    if (find === replace) {
      return { ok: false, error: `Edit ${index} must have different find and replace text.` };
    }

    const reason = typeof edit.reason === "string" ? edit.reason.trim() : undefined;
    if (reason && reason.length > MAX_REASON_LENGTH) {
      return { ok: false, error: `Edit ${index} reason exceeds ${MAX_REASON_LENGTH} characters.` };
    }

    const rawMatchCase = edit.matchCase;
    if (rawMatchCase !== undefined && typeof rawMatchCase !== "boolean") {
      return { ok: false, error: `Edit ${index} matchCase must be a boolean.` };
    }
    const rawMatchWholeWord = edit.matchWholeWord;
    if (rawMatchWholeWord !== undefined && typeof rawMatchWholeWord !== "boolean") {
      return { ok: false, error: `Edit ${index} matchWholeWord must be a boolean.` };
    }
    const matchCase = rawMatchCase === true;
    const matchWholeWord = rawMatchWholeWord === true;
    const searchKey = JSON.stringify([find, matchCase, matchWholeWord]);
    const previousIndex = seenSearches.get(searchKey);
    if (previousIndex !== undefined) {
      return {
        ok: false,
        error: `Edit ${index} duplicates edit ${previousIndex} for the same search text and match options.`
      };
    }
    seenSearches.set(searchKey, index);
    const highlightColor = normalizeApplyEditPlanHighlightColor(edit.highlightColor);
    if (edit.highlightColor !== undefined && highlightColor === null) {
      return { ok: false, error: `Edit ${index} uses an unsupported highlight color.` };
    }

    edits.push({
      find,
      replace,
      matchCase,
      matchWholeWord,
      reason: reason || undefined,
      ...(highlightColor ? { highlightColor } : {})
    });
  }

  return { ok: true, plan: { edits } };
}

async function applyEditPlan(args: JsonRecord, context?: Word.RequestContext): Promise<string> {
  const validated = validateApplyEditPlan(args);
  if (validated.ok === false) {
    return stringifyError(validated.error);
  }

  if (!Office.context?.requirements?.isSetSupported("WordApi", "1.3")) {
    return stringifyError("apply_edit_plan requires WordApi 1.3, which this Word host does not support.");
  }

  const requestedEdits = validated.plan.edits;
  const state: OperationState = { phase: "search", mutationState: "none" };

  return safelyWord(async (context) => {
    const body = context.document.body;
    const prepared = requestedEdits.map((edit, index) => {
      const matches = body.search(literalQuery(edit.find, index), literalOptions(edit.matchCase, edit.matchWholeWord));
      matches.load("items");
      return { edit, index, matches };
    });

    await context.sync();

    state.phase = "compare";
    const locationComparisons: Array<{
      leftIndex: number;
      rightIndex: number;
      result: OfficeExtension.ClientResult<Word.LocationRelation>;
    }> = [];
    for (let leftIndex = 0; leftIndex < prepared.length; leftIndex += 1) {
      const left = prepared[leftIndex];
      if (left.matches.items.length !== 1) {
        continue;
      }

      for (let rightIndex = leftIndex + 1; rightIndex < prepared.length; rightIndex += 1) {
        const right = prepared[rightIndex];
        if (right.matches.items.length !== 1) {
          continue;
        }
        locationComparisons.push({
          leftIndex,
          rightIndex,
          result: left.matches.items[0].compareLocationWith(right.matches.items[0])
        });
      }
    }
    if (locationComparisons.length > 0) {
      await context.sync();
    }

    const overlappingIndexes = new Set<number>();
    for (const comparison of locationComparisons) {
      if (!rangesOverlap(comparison.result.value)) {
        continue;
      }
      overlappingIndexes.add(comparison.leftIndex);
      overlappingIndexes.add(comparison.rightIndex);
    }

    const executions: ApplyEditPlanExecutionItem[] = [];

    for (const item of prepared) {
      const matchCount = item.matches.items.length;
      const isOverlapping = overlappingIndexes.has(item.index);
      const execution: ApplyEditPlanExecutionItem = {
        index: item.index,
        find: item.edit.find,
        replace: item.edit.replace,
        status: isOverlapping ? "overlapping" : matchCount === 1 ? "applied" : matchCount === 0 ? "not_found" : "ambiguous",
        matchCount,
        highlighted: Boolean(item.edit.highlightColor),
        reason: item.edit.reason ?? null
      };

      if (execution.status === "applied") {
        state.phase = "write";
        state.mutationState = "unknown";
        const match = item.matches.items[0];
        const inserted = match.insertText(item.edit.replace, "Replace");
        if (item.edit.highlightColor) {
          const normalized = normalizeApplyEditPlanHighlightColor(item.edit.highlightColor);
          if (normalized) {
            inserted.font.highlightColor = normalized;
          }
        }
      }

      executions.push(execution);
    }

    if (executions.some((item) => item.status === "applied")) {
      await context.sync();
    }

    let totalApplied = 0;
    let totalSkipped = 0;
    for (const execution of executions) {
      if (execution.status === "applied") {
        totalApplied += 1;
      } else {
        totalSkipped += 1;
      }
    }

    const result: ApplyEditPlanExecutionResult = {
      ok: true,
      action: "apply_edit_plan",
      totalRequested: requestedEdits.length,
      totalApplied,
      totalSkipped,
      edits: executions
    };

    return result;
  }, context, state) as Promise<string>;
}

export async function getDocumentContext(): Promise<string> {
  const snapshot = await getDocumentSnapshot(CONTEXT_READ_CHARS);
  return `Overview:\n${JSON.stringify(snapshot.overview)}\n\nFresh document text:\n${JSON.stringify(snapshot.documentText)}`;
}

/** Local-loop verification after an awaited mutation; independent fresh reads. */
export async function getPostMutationReadback() {
  const text = safeParseJson(await readDocumentText(MAX_READ_CHARS));
  const structure = safeParseJson(await getDocumentStructure(100));
  return { text, structure };
}

export function supportsWordApi14(): boolean {
  try {
    return !!(Office?.context?.requirements?.isSetSupported?.("WordApi", "1.4"));
  } catch {
    return false;
  }
}

/** Optional desktop-only range offsets; the manifest remains WordApi 1.2. */
export function supportsWordApiDesktop14(): boolean {
  try {
    return !!(Office?.context?.requirements?.isSetSupported?.("WordApiDesktop", "1.4"));
  } catch {
    return false;
  }
}

export async function getDocumentSnapshot(maxLength = CONTEXT_READ_CHARS): Promise<DocumentSnapshot> {
  const raw = await safelyWord(async (context) => {
    const hasDesktopOffsets = supportsWordApiDesktop14();
    const document = context.document;
    const body = document.body;
    const selection = document.getSelection();
    const paragraphs = body.paragraphs;

    body.load("text");
    paragraphs.load("items");
    if (hasDesktopOffsets) {
      selection.load(["text", "start", "end"]);
    } else {
      selection.load("text");
    }
    await context.sync();

    const words = body.text.split(/\s+/).filter(Boolean).length;

    return {
      overview: {
        charCount: body.text.length,
        wordCount: words,
        paragraphCount: paragraphs.items.length,
        selection: {
          start: hasDesktopOffsets && Number.isFinite(selection.start) ? selection.start : null,
          end: hasDesktopOffsets && Number.isFinite(selection.end) ? selection.end : null,
          text: truncateText(selection.text, 300).text
        }
      },
      documentText: truncateText(body.text, normalizeMaxLength(maxLength))
    };
  });

  const parsed = safeParseJson(raw);
  if (parsed && typeof parsed === "object" && "ok" in parsed && (parsed as { ok?: unknown }).ok === false) {
    const message = "error" in parsed ? String((parsed as { error?: unknown }).error ?? "Word snapshot failed.") : "Word snapshot failed.";
    throw new Error(message);
  }
  if (!parsed || typeof parsed !== "object" || !("overview" in parsed) || !("documentText" in parsed)) {
    throw new Error("Could not build document snapshot from Word API state.");
  }

  return parsed as DocumentSnapshot;
}

async function getDocumentOverview(context?: Word.RequestContext): Promise<string> {
  return safelyWord(async (context) => {
    const hasDesktopOffsets = supportsWordApiDesktop14();
    const document = context.document;
    const body = document.body;
    const selection = document.getSelection();
    const paragraphs = body.paragraphs;

    body.load("text");
    paragraphs.load("items");
    if (hasDesktopOffsets) {
      selection.load(["text", "start", "end"]);
    } else {
      selection.load("text");
    }
    await context.sync();

    const words = body.text.split(/\s+/).filter(Boolean).length;
    const paragraphCount = paragraphs.items.length;

    return {
      ok: true,
      action: "get_document_overview",
      charCount: body.text.length,
      wordCount: words,
      paragraphCount,
      selection: {
        start: hasDesktopOffsets && Number.isFinite(selection.start) ? selection.start : null,
        end: hasDesktopOffsets && Number.isFinite(selection.end) ? selection.end : null,
        text: truncateText(selection.text, 300).text
      }
    };
  }, context);
}

async function readSelection(context?: Word.RequestContext): Promise<string> {
  return safelyWord(async (context) => {
    const hasDesktopOffsets = supportsWordApiDesktop14();
    const selection = context.document.getSelection();
    if (hasDesktopOffsets) {
      selection.load(["text", "start", "end"]);
    } else {
      selection.load("text");
    }
    await context.sync();

    const truncated = truncateText(selection.text, MAX_READ_CHARS);
    return {
      ok: true,
      action: "read_selection",
      truncated: truncated.truncated,
      originalLength: truncated.originalLength,
      returnedLength: truncated.returnedLength,
      start: hasDesktopOffsets && Number.isFinite(selection.start) ? selection.start : null,
      end: hasDesktopOffsets && Number.isFinite(selection.end) ? selection.end : null,
      text: truncated.text
    };
  }, context);
}

async function readDocumentText(maxLength: number, context?: Word.RequestContext): Promise<string> {
  return safelyWord(async (context) => {
    const normalizedMaxLength = normalizeMaxLength(maxLength);
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const truncated = truncateText(body.text, normalizedMaxLength);
    return {
      ok: true,
      action: "read_document_text",
      truncated: truncated.truncated,
      originalLength: truncated.originalLength,
      returnedLength: truncated.returnedLength,
      text: truncated.text
    };
  }, context);
}

async function insertText(text: string, location: string, context?: Word.RequestContext): Promise<string> {
  if (text.length > MAX_REPLACE_LENGTH) {
    return stringifyError(new Error(`Argument "text" must not exceed ${MAX_REPLACE_LENGTH} characters.`));
  }

  const normalizedLocation = normalizeInsertionLocation(location, "insert_text");

  return safelyWord(async (context) => {
    const document = context.document;

    if (normalizedLocation === "selection") {
      const selection = document.getSelection();
      selection.insertText(text, "Replace");
      await context.sync();
      return {
        ok: true,
        action: "insert_text",
        location: "selection",
        insertedLength: text.length
      };
    }

    const body = document.body;
    body.insertText(text, normalizedLocation === "start" ? "Start" : "End");
    await context.sync();
    return {
      ok: true,
      action: "insert_text",
      location: normalizedLocation,
      insertedLength: text.length
    };
  }, context, { phase: "write", mutationState: "unknown" });
}

function normalizeInsertionLocation(rawLocation: string | undefined, toolName: string): "start" | "end" | "selection" {
  if (rawLocation === undefined || rawLocation === "") {
    return "end";
  }

  if (rawLocation === "start" || rawLocation === "end" || rawLocation === "selection") {
    return rawLocation;
  }

  throw new Error(`${toolName} received an invalid location. Use "start", "end", or "selection".`);
}

async function replaceSelection(text: string, highlightColor?: string, context?: Word.RequestContext): Promise<string> {
  if (text.length > MAX_REPLACE_LENGTH) {
    return stringifyError(new Error(`Argument "text" must not exceed ${MAX_REPLACE_LENGTH} characters.`));
  }

  const normalizedHighlight = normalizeAndValidateHighlightColor(highlightColor, "replace_selection");

  return safelyWord(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    const inserted = selection.insertText(text, "Replace");
    if (normalizedHighlight) {
      inserted.font.highlightColor = normalizedHighlight;
    }
    await context.sync();

    return {
      ok: true,
      action: "replace_selection",
      previousLength: selection.text.length,
      insertedLength: text.length,
      highlighted: Boolean(normalizedHighlight),
      highlightColor: normalizedHighlight || null
    };
  }, context, { phase: "write", mutationState: "unknown" });
}

async function findReplace(
  find: string,
  replacement: string,
  matchCase: boolean,
  matchWholeWord: boolean,
  highlightColor?: string,
  context?: Word.RequestContext
): Promise<string> {
  if (find === replacement) {
    return stringifyError(new Error(`Argument "find" and "replace" must be different.`));
  }

  if (find.length > MAX_FIND_LENGTH) {
    return stringifyError(new Error(`Argument "find" must not exceed ${MAX_FIND_LENGTH} characters.`));
  }

  if (replacement.length > MAX_REPLACE_LENGTH) {
    return stringifyError(new Error(`Argument "replace" must not exceed ${MAX_REPLACE_LENGTH} characters.`));
  }

  const normalizedHighlight = normalizeAndValidateHighlightColor(highlightColor, "find_replace");
  const query = literalQuery(find);
  const state: OperationState = { phase: "search", mutationState: "none" };

  return safelyWord(async (context) => {
    const matches = context.document.body.search(query, literalOptions(matchCase, matchWholeWord));
    matches.load("items");
    await context.sync();

    if (matches.items.length > 0) {
      state.phase = "write";
      state.mutationState = "unknown";
      for (const match of matches.items) {
        const inserted = match.insertText(replacement, "Replace");
        if (normalizedHighlight) {
          inserted.font.highlightColor = normalizedHighlight;
        }
      }

      await context.sync();
    }

    return {
      ok: true,
      action: "find_replace",
      find,
      replace: replacement,
      matchCase,
      matchWholeWord,
      replacements: matches.items.length,
      highlighted: Boolean(normalizedHighlight),
      highlightColor: normalizedHighlight || null
    };
  }, context, state);
}

async function formatSelection(
  bold: boolean | undefined,
  italic: boolean | undefined,
  color: string | undefined,
  highlightColor: string | undefined,
  fontSize: number | undefined,
  context?: Word.RequestContext
): Promise<string> {
  if (bold === undefined && italic === undefined && color === undefined && highlightColor === undefined && fontSize === undefined) {
    return stringifyError("format_selection requires at least one formatting property.");
  }

  const normalizedHighlight = normalizeAndValidateHighlightColor(highlightColor, "format_selection");

  return safelyWord(async (context) => {
    const selection = context.document.getSelection();
    const normalizedFontSize = fontSize;

    if (bold !== undefined) {
      selection.font.bold = bold;
    }
    if (italic !== undefined) {
      selection.font.italic = italic;
    }
    if (color !== undefined) {
      selection.font.color = color;
    }
    if (normalizedHighlight !== undefined) {
      selection.font.highlightColor = normalizedHighlight;
    }
    if (normalizedFontSize !== undefined) {
      selection.font.size = normalizedFontSize;
    }

    selection.load("text");
    await context.sync();

    const changedProperties = [bold, italic, color, highlightColor, fontSize]
      .filter((value) => value !== undefined).length;

    return {
      ok: true,
      action: "format_selection",
      changedProperties,
      selectedLength: selection.text.length
    };
  }, context, { phase: "write", mutationState: "unknown" });
}

async function getDocumentStructure(maxParagraphs?: number, context?: Word.RequestContext): Promise<string> {
  const builtInSupported = supportsBuiltInStyles();
  return safelyWord(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    body.load("text");
    paragraphs.load("items");
    await context.sync();

    const previewCount = Math.min(Math.max(maxParagraphs ?? 10, 1), 100, paragraphs.items.length);
    if (previewCount > 0) {
      for (const paragraph of paragraphs.items.slice(0, previewCount)) {
        paragraph.load(builtInSupported ? ["text", "style", "styleBuiltIn"] : ["text", "style"]);
      }
      await context.sync();
    }
    return {
      ok: true,
      action: "get_document_structure",
      styleBuiltInSupported: builtInSupported,
      totalParagraphs: paragraphs.items.length,
      totalCharacters: body.text.length,
      paragraphPreviews: paragraphs.items.slice(0, previewCount).map((paragraph, index) => {
        const text = paragraph.text ?? "";
        return { index: index + 1, text: text.length > 200 ? `${text.slice(0, 200)}...` : text, length: text.length, truncated: text.length > 200, ...observedStyle(paragraph, builtInSupported) };
      })
    };
  }, context);
}

function supportsBuiltInStyles(): boolean {
  try { return !!Office.context?.requirements?.isSetSupported("WordApi", "1.3"); } catch { return false; }
}

function observedStyle(paragraph: Word.Paragraph, builtInSupported: boolean) {
  return { style: paragraph.style, styleBuiltIn: builtInSupported ? paragraph.styleBuiltIn : null };
}

async function applyStyle(style: string, targetText?: string, context?: Word.RequestContext): Promise<string> {
  const builtInSupported = supportsBuiltInStyles();
  const builtIn = PARAGRAPH_STYLES.includes(style);
  if (builtIn && !builtInSupported) return JSON.stringify({ ok: false, error: "built_in_style_unsupported", phase: "validation", mutationState: "none", category: "office_host" });
  const query = targetText === undefined ? undefined : literalQuery(targetText);
  const state: OperationState = { phase: query === undefined ? "read" : "search", mutationState: "none" };
  return safelyWord(async (context) => {
    let paragraphs: Word.ParagraphCollection;
    if (query !== undefined) {
      const matches = context.document.body.search(query, literalOptions(true));
      matches.load("items"); await context.sync();
      if (matches.items.length !== 1) return { ok: false, error: "paragraph_target_not_unique", ...state };
      paragraphs = matches.items[0].paragraphs;
    } else paragraphs = context.document.getSelection().paragraphs;
    state.phase = "compare";
    paragraphs.load("items");
    await context.sync();
    if (targetText !== undefined) {
      if (paragraphs.items.length !== 1) return { ok: false, error: "paragraph_target_not_unique", ...state };
      paragraphs.items[0].load("text"); await context.sync();
      if (paragraphs.items[0].text !== targetText) return { ok: false, error: "paragraph_target_not_exact", ...state };
    }
    if (paragraphs.items.length > 0) {
      state.phase = "write"; state.mutationState = "unknown";
      for (const paragraph of paragraphs.items) {
        if (builtIn) paragraph.styleBuiltIn = style as Word.BuiltInStyleName;
        else paragraph.style = style;
      }
      await context.sync();
      state.phase = "readback";
      for (const paragraph of paragraphs.items) paragraph.load(builtInSupported ? ["style", "styleBuiltIn"] : ["style"]);
      await context.sync();
    }
    return { ok: true, action: "apply_style", style, styledParagraphs: paragraphs.items.length,
      styleBuiltInSupported: builtInSupported, observedStyles: paragraphs.items.map(p => observedStyle(p, builtInSupported)) };
  }, context, state);
}

async function insertParagraphs(args: JsonRecord, context?: Word.RequestContext): Promise<string> {
  const invalid = () => JSON.stringify({ ok: false, error: "invalid_paragraph_plan", category: "validation", phase: "validation", mutationState: "none" });
  const items = args.paragraphs;
  if (args.text !== undefined || !Array.isArray(items) || items.length < 1 || items.length > 20 || (args.location !== undefined && args.location !== "start" && args.location !== "end")) return invalid();
  let length = 0;
  for (const item of items) {
    if (!item || typeof item.text !== "string" || !item.text.trim() || /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/.test(item.text) || !PARAGRAPH_STYLES.includes(item.styleBuiltIn)) return invalid();
    length += item.text.length;
  }
  if (length > MAX_REPLACE_LENGTH) return invalid();
  if (!supportsBuiltInStyles()) return JSON.stringify({ ok: false, error: "built_in_style_unsupported", category: "office_host", phase: "validation", mutationState: "none" });
  const state: OperationState = { phase: "read", mutationState: "none" };
  return safelyWord(async ctx => {
    const body = ctx.document.body;
    const existing = body.paragraphs; existing.load("items"); await ctx.sync();
    if (existing.items.length === 1) { existing.items[0].load("text"); await ctx.sync(); }
    const empty = existing.items.length === 1 && existing.items[0].text === "";
    const created: Word.Paragraph[] = [];
    state.phase = "write"; state.mutationState = "unknown";
    for (const [index, item] of items.entries()) {
      let p: Word.Paragraph;
      // Text-empty is not proof of structural emptiness. Insert, never replace:
      // opaque fields, pictures and controls must remain owned by Word.
      if (index === 0 && empty) { p = existing.items[0]; p.insertText(item.text, args.location === "start" ? "Start" : "End"); }
      else if (index === 0) p = body.insertParagraph(item.text, args.location === "start" ? "Start" : "End");
      else p = created[index - 1].insertParagraph(item.text, "After");
      p.styleBuiltIn = item.styleBuiltIn as Word.BuiltInStyleName;
      created.push(p);
    }
    await ctx.sync(); state.phase = "readback";
    for (const p of created) p.load(["text", "style", "styleBuiltIn"]);
    await ctx.sync();
    return { ok: true, action: "insert_text", paragraphCount: created.length,
      observedParagraphs: created.map(p => ({ text: p.text, ...observedStyle(p, true) })) };
  }, context, state);
}

async function insertOoxml(ooxml: string, location: string, context?: Word.RequestContext): Promise<string> {
  if (ooxml.length > MAX_OOXML_LENGTH) {
    return stringifyError(new Error(`Argument "ooxml" must not exceed ${MAX_OOXML_LENGTH} characters.`));
  }

  const normalizedLocation = normalizeInsertionLocation(location, "insert_ooxml");

  return safelyWord(async (context) => {
    if (normalizedLocation === "selection") {
      context.document.getSelection().insertOoxml(ooxml, "Replace");
    } else {
      context.document.body.insertOoxml(ooxml, normalizedLocation === "start" ? "Start" : "End");
    }
    await context.sync();
    return { ok: true, action: "insert_ooxml", location: normalizedLocation, insertedLength: ooxml.length };
  }, context, { phase: "write", mutationState: "unknown" });
}

async function insertComment(text: string, context?: Word.RequestContext): Promise<string> {
  if (text.length > MAX_REPLACE_LENGTH) {
    return stringifyError(new Error(`Argument "text" must not exceed ${MAX_REPLACE_LENGTH} characters.`));
  }

  if (!supportsWordApi14()) {
    return stringifyError("insert_comment requires WordApi 1.4, which this Word host does not support.");
  }

  return safelyWord(async (context) => {
    const selection = context.document.getSelection();
    selection.insertComment(text);
    await context.sync();
    return { ok: true, action: "insert_comment", commentLength: text.length };
  }, context, { phase: "write", mutationState: "unknown" });
}

function normalizeMaxLength(raw: number): number {
  if (!Number.isFinite(raw) || raw < 1) {
    return MAX_READ_CHARS;
  }
  return Math.min(Math.max(Math.floor(raw), 1), 20000);
}

function safeParseJson(value: string): unknown | null {
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

export function truncateText(value: string, maxLength: number): ReadResult {
  const original = value ?? "";
  const text = original.slice(0, Math.max(0, Math.min(normalizeMaxLength(maxLength), original.length)));
  const returnedLength = text.length;
  return {
    text,
    truncated: returnedLength < original.length,
    originalLength: original.length,
    returnedLength
  };
}

function textArg(args: JsonRecord, name: string): string {
  const value = args[name];
  if (typeof value !== "string") {
    throw new Error(`Missing or invalid string argument "${name}".`);
  }
  return value;
}

function nonEmptyTextArg(args: JsonRecord, name: string): string {
  const value = textArg(args, name);
  if (!value.trim()) {
    throw new Error(`Argument "${name}" must not be empty.`);
  }
  return value;
}

function optionalTextArg(args: JsonRecord, name: string): string | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new Error(`Invalid string argument "${name}".`);
  }
  const trimmed = value.trim();
  return trimmed.length ? trimmed : undefined;
}

function optionalBooleanArg(args: JsonRecord, name: string): boolean | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "boolean") {
    throw new Error(`Invalid boolean argument "${name}".`);
  }
  return value;
}

function optionalNumberArg(args: JsonRecord, name: string): number | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`Invalid number argument "${name}".`);
  }
  return Math.floor(value);
}

function optionalBoundedPositiveIntegerArg(args: JsonRecord, name: string, maximum: number): number | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1 || value > maximum) {
    throw new Error(`Argument "${name}" must be an integer between 1 and ${maximum}.`);
  }
  return value;
}

function optionalMinimumNumberArg(args: JsonRecord, name: string, minimum: number): number | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`Invalid number argument "${name}".`);
  }
  if (value < minimum) {
    throw new Error(`Argument "${name}" must be at least ${minimum}.`);
  }
  return value;
}

function optionalPositiveIntegerArg(args: JsonRecord, name: string): number | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
    throw new Error(`Argument "${name}" must be a positive integer.`);
  }
  return value;
}

function normalizeApplyEditPlanHighlightColor(value: unknown): string | null | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const lower = trimmed.toLowerCase();
  if (lower === "yellow") {
    return "yellow";
  }

  if (/^#[0-9a-f]{6}$/i.test(trimmed)) {
    return trimmed.toUpperCase();
  }

  return null;
}

function normalizeAndValidateHighlightColor(value: string | undefined, toolName: string): string | undefined {
  const normalized = normalizeApplyEditPlanHighlightColor(value);
  if (normalized === null) {
    throw new Error(
      `${toolName} received an invalid highlightColor. Use "yellow" or a 6-digit hex value like "#FFFF00".`
    );
  }

  if (normalized === "yellow") {
    return "#FFFF00";
  }

  return normalized;
}

const OVERLAPPING_LOCATION_RELATIONS = new Set([
  "Equal",
  "ContainsStart",
  "ContainsEnd",
  "Contains",
  "InsideStart",
  "InsideEnd",
  "Inside",
  "OverlapsBefore",
  "OverlapsAfter"
]);

function rangesOverlap(relation: Word.LocationRelation): boolean {
  return OVERLAPPING_LOCATION_RELATIONS.has(String(relation));
}

async function safelyWord<T extends object>(
  work: (context: Word.RequestContext) => Promise<T>,
  context?: Word.RequestContext,
  state: OperationState = { phase: "read", mutationState: "none" }
): Promise<string> {
  try {
    const result = context ? await work(context) : await Word.run(work);
    return JSON.stringify(result);
  } catch (error) {
    return diagnosticFailure(error, state);
  }
}

function stringifyError(error: unknown, category?: ToolErrorCategory): string {
  if (error instanceof Error) {
    return JSON.stringify({ ok: false, error: error.message, name: error.name, phase: "validation", mutationState: "none", ...(category ? { category } : {}) });
  }
  return JSON.stringify({ ok: false, error: String(error), phase: "validation", mutationState: "none", ...(category ? { category } : {}) });
}
