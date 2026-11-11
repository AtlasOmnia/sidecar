import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const scriptDirectory = fileURLToPath(new URL(".", import.meta.url));
const repositoryRoot = resolve(scriptDirectory, "..");
const argumentsList = process.argv.slice(2);
const options = {
  check: false,
  contract: undefined,
  generated: undefined,
};

for (let index = 0; index < argumentsList.length; index += 1) {
  const argument = argumentsList[index];
  if (argument === "--check") {
    if (options.check) throw new Error("Duplicate --check argument.");
    options.check = true;
    continue;
  }
  if (argument === "--contract" || argument === "--generated") {
    const optionName = argument.slice(2);
    if (options[optionName] !== undefined) {
      throw new Error(`Duplicate ${argument} argument.`);
    }
    const value = argumentsList[index + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`${argument} requires a path argument.`);
    }
    options[optionName] = resolve(process.cwd(), value);
    index += 1;
    continue;
  }
  throw new Error(`Unknown argument: ${argument}`);
}

const contractPath = options.contract ?? resolve(repositoryRoot, "src/workbookPlanContract.ts");
const generatedPath = options.generated ?? resolve(repositoryRoot, "src/generated/workbookPlanSchema.ts");

const contractModule = createRequire(import.meta.url)(contractPath);
const contract = contractModule.workbookPlanContract;
if (!contract || typeof contract !== "object" || !contract.schema) {
  throw new Error("WorkbookPlan contract does not export a schema.");
}

const canonicalize = (value) => {
  if (Array.isArray(value)) {
    return value.map(canonicalize);
  }
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, item]) => [key, canonicalize(item)]),
    );
  }
  return value;
};

const schema = contract.schema;
const canonicalSchema = JSON.stringify(canonicalize(schema));
const fingerprint = createHash("sha256").update(canonicalSchema, "utf8").digest("hex");

const quote = (value) => JSON.stringify(value);

const typeFor = (node) => {
  if (node.$ref) {
    return node.$ref.replace(/^#\/\$defs\//, "");
  }
  if (node.const !== undefined) {
    return quote(node.const);
  }
  if (node.enum) {
    return node.enum.map(quote).join(" | ");
  }
  if (node.oneOf) {
    return node.oneOf.map(typeFor).join(" | ");
  }
  if (node.type === "object") {
    const properties = node.properties ?? {};
    const required = new Set(node.required ?? []);
    const members = Object.entries(properties).map(([key, property]) => {
      const optional = required.has(key) ? "" : "?";
      return `readonly ${key}${optional}: ${typeFor(property)};`;
    });
    return `{ ${members.join(" ")} }`;
  }
  if (node.type === "array") {
    return `ReadonlyArray<${typeFor(node.items ?? {})}>`;
  }
  if (node.type === "string") {
    return "string";
  }
  if (node.type === "integer" || node.type === "number") {
    return "number";
  }
  if (node.type === "boolean") {
    return "boolean";
  }
  if (node.type === "null") {
    return "null";
  }
  return "unknown";
};

const definitionTypes = Object.entries(schema.$defs ?? {})
  .map(([name, definition]) => `export type ${name} = ${typeFor(definition)};`)
  .join("\n");

const output = [
  "/* Generated from the repository-authored WorkbookPlan V1 contract. Do not edit. */",
  `export const workbookPlanSchema = ${JSON.stringify(schema, null, 2)} as const;`,
  "",
  `export const workbookPlanFingerprint = ${quote(fingerprint)} as const;`,
  "export const WORKBOOK_PLAN_FINGERPRINT = workbookPlanFingerprint;",
  "",
  definitionTypes,
  "",
  `export type WorkbookPlan = ${typeFor(schema)};`,
  "",
].join("\n");

const checkOnly = process.argv.includes("--check");
if (checkOnly) {
  let existing;
  try {
    existing = await readFile(generatedPath, "utf8");
  } catch {
    console.error("Generated WorkbookPlan output is missing; run npm run generate:workbook-plan.");
    process.exitCode = 1;
  }
  if (existing !== undefined && existing !== output) {
    console.error("Generated WorkbookPlan output is out of date; run npm run generate:workbook-plan.");
    process.exitCode = 1;
  }
} else {
  await mkdir(resolve(generatedPath, ".."), { recursive: true });
  await writeFile(generatedPath, output, "utf8");
  console.log("Generated WorkbookPlan schema and types.");
}
