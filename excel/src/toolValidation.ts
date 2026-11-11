/**
 * Strict pre-execution tool-call validation.
 *
 * Checks performed before touching Office.js:
 * 1. Tool name is known
 * 2. Arguments are valid JSON
 * 3. Required fields are present
 * 4. All supplied fields are declared in the schema (no unknown fields)
 * 5. Field types match the schema (basic type checks)
 */

import { OpenAITool } from "./types";
import { ToolError, ToolErrorCategory, validationError } from "./toolErrors";
import { excelTools } from "./excelTools";

const toolMap = new Map<string, { schema: OpenAITool; index: number }>();
for (let i = 0; i < excelTools.length; i++) {
  toolMap.set(excelTools[i].schema.function.name, {
    schema: excelTools[i].schema,
    index: i,
  });
}

export interface ValidationResult {
  valid: true;
  toolIndex: number;
  parsedArgs: Record<string, unknown>;
}

/**
 * Validate a tool call before execution. Returns either a structured error
 * or a success with the parsed args and tool index.
 */
export function validateToolCall(
  name: string,
  rawArgs: string,
): ToolError | ValidationResult {
  const entry = toolMap.get(name);
  if (!entry) {
    return validationError(
      `Unknown tool "${name}". Check available tools and try again.`,
    );
  }

  const schema = entry.schema;
  const params = schema.function.parameters as Record<string, unknown> | undefined;
  const properties = (params?.properties ?? {}) as Record<string, unknown>;
  const required = (params?.required ?? []) as string[];
  const additionalAllowed =
    params?.additionalProperties !== false;

  // Parse JSON
  let args: Record<string, unknown>;
  try {
    args = rawArgs.trim() ? (JSON.parse(rawArgs) as Record<string, unknown>) : {};
  } catch {
    return validationError(
      `Tool arguments for "${name}" are not valid JSON.`,
    );
  }

  if (typeof args !== "object" || args === null || Array.isArray(args)) {
    return validationError(
      `Tool arguments for "${name}" must be a JSON object.`,
    );
  }

  // Check required fields
  for (const req of required) {
    if (!(req in args)) {
      return validationError(
        `Missing required argument "${req}" for tool "${name}".`,
      );
    }
  }

  // Check unknown fields
  if (!additionalAllowed) {
    for (const key of Object.keys(args)) {
      if (!(key in properties)) {
        return validationError(
          `Unknown argument "${key}" for tool "${name}". Allowed fields: ${Object.keys(properties).join(", ") || "(none)"}.`,
        );
      }
    }
  }

  // Basic type checks on supplied fields
  for (const [key, value] of Object.entries(args)) {
    const propSchema = properties[key] as Record<string, unknown> | undefined;
    if (!propSchema) continue; // already caught above if additionalProperties false

    const typeError = checkFieldType(key, value, propSchema);
    if (typeError) return typeError;
  }

  return {
    valid: true,
    toolIndex: entry.index,
    parsedArgs: args,
  };
}

function checkFieldType(
  key: string,
  value: unknown,
  propSchema: Record<string, unknown>,
): ToolError | null {
  const type = propSchema.type as string | undefined;
  if (!type) return null;

  switch (type) {
    case "string":
      if (typeof value !== "string") {
        return validationError(
          `Argument "${key}" must be a string, got ${jsType(value)}.`,
        );
      }
      break;
    case "number":
    case "integer":
      if (typeof value !== "number") {
        return validationError(
          `Argument "${key}" must be a number, got ${jsType(value)}.`,
        );
      }
      if (type === "integer" && !Number.isInteger(value)) {
        return validationError(
          `Argument "${key}" must be an integer, got ${value}.`,
        );
      }
      break;
    case "boolean":
      if (typeof value !== "boolean") {
        return validationError(
          `Argument "${key}" must be a boolean, got ${jsType(value)}.`,
        );
      }
      break;
    case "array":
      if (!Array.isArray(value)) {
        return validationError(
          `Argument "${key}" must be an array, got ${jsType(value)}.`,
        );
      }
      break;
    case "object":
      if (typeof value !== "object" || value === null || Array.isArray(value)) {
        return validationError(
          `Argument "${key}" must be an object, got ${jsType(value)}.`,
        );
      }
      break;
    default:
      // Unknown type — skip strict check but warn in development
      break;
  }

  // Enum check
  if (propSchema.enum && Array.isArray(propSchema.enum)) {
    if (!propSchema.enum.includes(value)) {
      return validationError(
        `Argument "${key}" must be one of [${propSchema.enum.join(", ")}], got ${JSON.stringify(value)}.`,
      );
    }
  }

  return null;
}

function jsType(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  return typeof value;
}

/**
 * Parse tool arguments without full validation — used where validation
 * is handled by the caller (e.g., tool loop with pre-validated calls).
 */
export function parseToolArgs(rawArgs: string): Record<string, unknown> | null {
  try {
    const parsed = rawArgs.trim() ? JSON.parse(rawArgs) : {};
    if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
      return null;
    }
    return parsed as Record<string, unknown>;
  } catch {
    return null;
  }
}
