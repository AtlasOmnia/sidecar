import { toolSchemas } from "./powerPointTools";
import { ToolError, ToolErrorCategory, validationError } from "./toolErrors";

export type JsonObject = Record<string, unknown>;
export type ValidationSuccess = { valid: true; toolIndex: number; parsedArgs: JsonObject };
export type ValidationResult = ValidationSuccess | (ToolError & { valid: false });

export function parseToolArgs(raw: string): JsonObject | null {
  if (!raw.trim()) return {};
  try {
    const value: unknown = JSON.parse(raw);
    return value !== null && typeof value === "object" && !Array.isArray(value) ? value as JsonObject : null;
  } catch { return null; }
}

export function validateToolCall(name: string, raw: string): ValidationResult {
  const toolIndex = toolSchemas.findIndex((tool) => tool.function.name === name);
  if (toolIndex < 0) return fail(`Unknown tool: ${name}`);
  const args = parseToolArgs(raw);
  if (args === null) return fail("Tool arguments must be a JSON object.");
  const schema = toolSchemas[toolIndex].function.parameters;
  const result = validateSchemaValue(args, schema, "");
  if (result) return fail(result);
  return { valid: true, toolIndex, parsedArgs: args };
}

export function validateSchemaValue(value: unknown, schema: Record<string, unknown>, path = ""): string | null {
  const label = path || "arguments";
  const type = schema.type;
  if (type === "object") {
    if (value === null || typeof value !== "object" || Array.isArray(value)) return `${label} must be an object.`;
    const object = value as Record<string, unknown>;
    const properties = (schema.properties && typeof schema.properties === "object" ? schema.properties : {}) as Record<string, Record<string, unknown>>;
    const required = Array.isArray(schema.required) ? schema.required.filter((item): item is string => typeof item === "string") : [];
    for (const key of required) if (!(key in object)) return `Missing or invalid required argument: ${path ? `${path}.` : ""}${key}`;
    if (schema.additionalProperties === false) for (const key of Object.keys(object)) if (!(key in properties)) return `Unknown argument: ${path ? `${path}.` : ""}${key}`;
    for (const [key, child] of Object.entries(properties)) if (key in object) { const error = validateSchemaValue(object[key], child, path ? `${path}.${key}` : key); if (error) return error; }
    return null;
  }
  if (type === "array") {
    if (!Array.isArray(value)) return `${label} must be an array.`;
    if (typeof schema.minItems === "number" && value.length < schema.minItems) return `${label} must contain at least ${schema.minItems} item(s).`;
    if (typeof schema.maxItems === "number" && value.length > schema.maxItems) return `${label} must contain at most ${schema.maxItems} item(s).`;
    const itemSchema = schema.items && typeof schema.items === "object" ? schema.items as Record<string, unknown> : null;
    if (itemSchema) for (let i = 0; i < value.length; i += 1) { const error = validateSchemaValue(value[i], itemSchema, `${label}[${i}]`); if (error) return error; }
    return null;
  }
  if (type === "string") { if (typeof value !== "string") return `${label} must be a string argument.`; if (typeof schema.minLength === "number" && value.length < schema.minLength) return `${label} must have at least ${schema.minLength} characters.`; }
  else if (type === "number") { if (typeof value !== "number" || !Number.isFinite(value)) return `${label} must be a finite number.`; if (schema.integer === true && !Number.isInteger(value)) return `${label} must be an integer.`; if (schema.minimum !== undefined && value < Number(schema.minimum)) return `${label} is below minimum ${schema.minimum}.`; if (schema.maximum !== undefined && value > Number(schema.maximum)) return `${label} exceeds maximum ${schema.maximum}.`; }
  else if (type === "integer") {
    if (typeof value !== "number" || !Number.isFinite(value) || !Number.isInteger(value)) return `${label} must be an integer.`;
    if (schema.minimum !== undefined && value < Number(schema.minimum)) return `${label} is below minimum ${schema.minimum}.`;
    if (schema.maximum !== undefined && value > Number(schema.maximum)) return `${label} exceeds maximum ${schema.maximum}.`;
  }
  else if (type === "boolean") { if (typeof value !== "boolean") return `${label} must be a boolean.`; }
  if (Array.isArray(schema.enum) && !schema.enum.some((candidate) => Object.is(candidate, value))) return `${label} must be one of: ${schema.enum.join(", ")}.`;
  return null;
}

function fail(message: string): ValidationResult { return Object.assign(validationError(message), { valid: false, category: ToolErrorCategory.Validation }); }


export function validateToolArguments(rawArgs: string, schema?: Record<string, unknown>, toolName?: string): { ok: true; parsedArgs: JsonObject } | { ok: false; error: ToolError } {
  if (!schema) return { ok: false, error: validationError(`Unknown tool: ${toolName || "unknown"}`) };
  let value: unknown;
  try { value = rawArgs.trim() ? JSON.parse(rawArgs) : {}; } catch { return { ok: false, error: validationError("Invalid JSON arguments.") }; }
  if (value === null || typeof value !== "object" || Array.isArray(value)) return { ok: false, error: validationError("arguments must be an object.") };
  const error = validateSchemaValue(value, schema, "arguments");
  if (error) return { ok: false, error: validationError(error) };
  return { ok: true, parsedArgs: value as JsonObject };
}
