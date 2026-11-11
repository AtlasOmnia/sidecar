/** Shared schema parsing and validation used by every pre-execution boundary. */

export interface ParsedToolArguments {
  valid: true;
  parsedArgs: Record<string, unknown>;
}

export function parseAndValidateToolArguments(
  rawArgs: string,
  schema: Record<string, unknown>,
): string | ParsedToolArguments {
  let args: unknown;
  try {
    args = rawArgs.trim() ? JSON.parse(rawArgs) : {};
  } catch {
    return "are not valid JSON.";
  }

  const schemaError = validateSchema(args, schema, "");
  if (schemaError) return schemaError;
  return { valid: true, parsedArgs: args as Record<string, unknown> };
}

export function validateSchema(value: unknown, schema: Record<string, unknown>, path: string): string | null {
  const type = typeof schema.type === "string" ? schema.type : undefined;
  const label = path || "arguments";

  if (Array.isArray(schema.enum) && !schema.enum.some((item) => Object.is(item, value))) {
    return `${label} must be one of [${schema.enum.join(", ")}].`;
  }

  if (type === "object") {
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
      return `${label} must be an object, got ${jsType(value)}.`;
    }
    const objectValue = value as Record<string, unknown>;
    const properties = (schema.properties ?? {}) as Record<string, unknown>;
    const required = Array.isArray(schema.required) ? schema.required as string[] : [];
    for (const key of required) {
      if (!Object.prototype.hasOwnProperty.call(objectValue, key)) {
        return `is missing required field ${pathFor(path, key)}.`;
      }
    }
    if (typeof schema.minProperties === "number" && Object.keys(objectValue).length < schema.minProperties) {
      return `${label} must contain at least ${schema.minProperties} properties.`;
    }
    if (schema.additionalProperties === false) {
      for (const key of Object.keys(objectValue)) {
        if (!Object.prototype.hasOwnProperty.call(properties, key)) {
          return `Unknown field ${pathFor(path, key)}.`;
        }
      }
    }
    for (const [key, child] of Object.entries(properties)) {
      if (Object.prototype.hasOwnProperty.call(objectValue, key)) {
        const error = validateSchema(objectValue[key], child as Record<string, unknown>, pathFor(path, key));
        if (error) return error;
      }
    }
  } else if (type === "array") {
    if (!Array.isArray(value)) return `${label} must be an array, got ${jsType(value)}.`;
    if (typeof schema.minItems === "number" && value.length < schema.minItems) return `${label} must contain at least ${schema.minItems} items.`;
    if (typeof schema.maxItems === "number" && value.length > schema.maxItems) return `${label} must contain at most ${schema.maxItems} items.`;
    if (schema.items && typeof schema.items === "object") {
      for (let index = 0; index < value.length; index += 1) {
        const error = validateSchema(value[index], schema.items as Record<string, unknown>, `${path}[${index}]`);
        if (error) return error;
      }
    }
  } else if (type === "string") {
    if (typeof value !== "string") return `${label} must be a string, got ${jsType(value)}.`;
    if (typeof schema.minLength === "number" && value.length < schema.minLength) return `${label} must be at least ${schema.minLength} characters.`;
    if (typeof schema.maxLength === "number" && value.length > schema.maxLength) return `${label} must not exceed ${schema.maxLength} characters.`;
  } else if (type === "number" || type === "integer") {
    if (typeof value !== "number" || !Number.isFinite(value)) return `${label} must be a number, got ${jsType(value)}.`;
    if (type === "integer" && !Number.isInteger(value)) return `${label} must be an integer, got ${value}.`;
    if (typeof schema.minimum === "number" && value < schema.minimum) return `${label} must be at least ${schema.minimum}.`;
    if (typeof schema.maximum === "number" && value > schema.maximum) return `${label} must be at most ${schema.maximum}.`;
  } else if (type === "boolean" && typeof value !== "boolean") {
    return `${label} must be a boolean, got ${jsType(value)}.`;
  }

  return null;
}

function pathFor(parent: string, key: string): string {
  return parent ? `${parent}.${key}` : key;
}

function jsType(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  return typeof value;
}
