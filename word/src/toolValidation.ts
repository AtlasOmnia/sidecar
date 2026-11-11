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
import { ToolError, validationError } from "./toolErrors";
import { validateSchema } from "./toolValidationCore";
import { wordTools } from "./wordTools";

const toolMap = new Map<string, { schema: OpenAITool; index: number }>();
for (let i = 0; i < wordTools.length; i++) {
  toolMap.set(wordTools[i].schema.function.name, {
    schema: wordTools[i].schema,
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

  let args: unknown;
  try {
    args = rawArgs.trim() ? JSON.parse(rawArgs) : {};
  } catch {
    return validationError(`Tool arguments for "${name}" are not valid JSON.`);
  }

  const params = (entry.schema.function.parameters ?? { type: "object" }) as Record<string, unknown>;
  const schemaError = validateSchema(args, params, "");
  if (schemaError) return validationError(`Tool "${name}" ${schemaError}`);

  return {
    valid: true,
    toolIndex: entry.index,
    parsedArgs: args as Record<string, unknown>,
  };
}
