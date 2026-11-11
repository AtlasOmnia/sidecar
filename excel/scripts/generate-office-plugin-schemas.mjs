#!/usr/bin/env node
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { lstatSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, readlinkSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(scriptDir, "..");
const options = parseArguments(process.argv.slice(2));
const outputPath = options.output ?? join(repoRoot, "plugins", "atlas-sidecar-excel", "tool_schemas.json");
const skillPath = options.skill ?? join(repoRoot, "plugins", "atlas-sidecar-excel", "skills", "workbook-plans", "SKILL.md");
const checkOnly = options.check;

function parseArguments(args) {
  const options = { check: false };
  const seen = new Set();
  for (let index = 0; index < args.length; index += 1) {
    const flag = args[index];
    if (!["--check", "--output", "--skill"].includes(flag)) throw new Error(`Unknown argument: ${flag}`);
    if (seen.has(flag)) throw new Error(`Duplicate flag: ${flag}`);
    seen.add(flag);
    if (flag === "--check") {
      options.check = true;
    } else {
      const value = args[++index];
      if (!value || value.startsWith("--")) throw new Error(`Missing value for ${flag}`);
      options[flag.slice(2)] = resolve(value);
    }
  }
  return options;
}

// Resolve existing symlinks, including ancestors of not-yet-created targets.
// Fail closed for permissions, loops and non-directory ancestors.
function realTarget(path) {
  return resolveTarget(path, new Set());
}

function resolveTarget(path, seen) {
  if (seen.has(path)) throw new Error("Output/input path contains a symlink cycle");
  seen.add(path);
  try {
    return realpathSync(path);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    // realpath also returns ENOENT for a dangling symlink: follow its target
    // rather than treating the symlink's own name as a new regular file.
    if (lstatSync(path, { throwIfNoEntry: false })?.isSymbolicLink()) {
      return resolveTarget(resolve(dirname(path), readlinkSync(path)), seen);
    }
    const parent = dirname(path);
    if (parent === path) throw error;
    return join(resolveTarget(parent, seen), basename(path));
  }
}

function sourceFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    return entry.isDirectory() ? sourceFiles(path) : [path];
  });
}

function refuseOutputAliases() {
  const targets = [outputPath, skillPath].map(realTarget);
  if (targets[0] === targets[1]) throw new Error("Output and skill targets alias; they must name different files");
  // Protect the generated schema and every transitive local compilation source,
  // plus the generator itself and the compiler entry point it invokes.
  const inputs = new Set([
    ...sourceFiles(join(repoRoot, "src")),
    fileURLToPath(import.meta.url),
    join(repoRoot, "node_modules", "typescript", "bin", "tsc"),
  ].map(realTarget));
  if (targets.some((target) => inputs.has(target))) throw new Error("Output target aliases an input/source file");
}
refuseOutputAliases();
const tempDir = mkdtempSync(join(repoRoot, ".schema-build-"));

function compileCanonicalSchemas() {
  const tscCli = join(repoRoot, "node_modules", "typescript", "bin", "tsc");
  execFileSync(process.execPath, [
    tscCli,
    "src/excelTools.ts",
    "src/types.ts",
    "--target", "ES2021",
    "--module", "commonjs",
    "--moduleResolution", "node",
    "--esModuleInterop",
    "--skipLibCheck",
    "--types", "office-js",
    "--outDir", tempDir,
  ], { cwd: repoRoot, stdio: "pipe" });

  const require = createRequire(import.meta.url);
  const compiled = require(join(tempDir, "excelTools.js"));
  if (!Array.isArray(compiled.toolSchemas)) {
    throw new Error("Canonical src/excelTools.ts did not export toolSchemas");
  }
  return compiled.toolSchemas;
}

function enforceObjectContracts(value) {
  if (Array.isArray(value)) {
    value.forEach(enforceObjectContracts);
    return;
  }
  if (!value || typeof value !== "object") return;
  if (value.type === "object") value.additionalProperties = false;
  Object.values(value).forEach(enforceObjectContracts);
}

function buildPluginSchemas(canonicalSchemas) {
  return canonicalSchemas.map((wrapper) => {
    if (wrapper?.type !== "function" || !wrapper.function) {
      throw new Error("Unexpected canonical OpenAI tool schema shape");
    }
    const schema = structuredClone(wrapper.function);
    schema.name = `sidecar_excel_${schema.name}`;
    schema.parameters ??= { type: "object", properties: {} };
    schema.parameters.properties ??= {};
    enforceObjectContracts(schema.parameters);
    return schema;
  });
}

function buildWorkbookPlanSchemaTool() {
  const require = createRequire(import.meta.url);
  const generated = require(join(repoRoot, "src", "generated", "workbookPlanSchema.ts"));
  const parameters = generated?.workbookPlanSchema;
  const fingerprint = generated?.WORKBOOK_PLAN_FINGERPRINT;

  if (!parameters || typeof parameters !== "object") {
    throw new Error("Workbook plan generator output is missing a valid schema payload.");
  }
  if (typeof fingerprint !== "string" || !/^[0-9a-f]{64}$/.test(fingerprint)) {
    throw new Error("Workbook plan fingerprint was missing or malformed.");
  }

  const schema = {
    ...structuredClone(generated.workbookPlanSchema),
  };
  const tool = {
    type: "function",
    function: {
      name: "sidecar_excel_execute_plan",
      description:
        "Submit one bounded WorkbookPlan for deterministic inspect-or-execute handling. " +
        "Use one static function call only in plan mode.",
      parameters: schema,
    },
    "x-workbook-plan-fingerprint": fingerprint,
  };
  enforceObjectContracts(tool.function.parameters);
  return {
    ...tool.function,
    "x-workbook-plan-fingerprint": tool["x-workbook-plan-fingerprint"],
  };
}

// Fixed behavior sentences are pinned against the current kernel by the guidance tests.
// Only the schema/fingerprint generator owns contract data; do not create a second schema here.
function buildSkill(workbookPlan) {
  const schema = workbookPlan.parameters;
  return [
    "---",
    "name: workbook-plans",
    "description: Submit bounded WorkbookPlan V1 inspection and execution through the static Excel plan tool.",
    "---",
    "",
    "# Workbook plans",
    "",
    `At most one plan call per response. When submitting a plan, use the static \`${workbookPlan.name}\` tool. Do not mix plans with primitive tool calls.`,
    "After an execute plan call returns, or after an already_executed / uncertain reply, make no further plan calls in that turn; report the outcome instead. An inspect may be followed by at most one execute for the same turn.",
    `Allowed top-level fields (additionalProperties: ${schema.additionalProperties}): ${Object.keys(schema.properties).map((name) => `\`${name}\``).join(", ")}.`,
    `Outcomes: ${schema.properties.outcome.enum.join(", ")}.`,
    `Maximum operations: ${schema.properties.operations.maxItems}.`,
    "",
    "- Server inspection is supported when the plan rollout is enabled; rollout is default off.",
    "- When workbook facts are needed, emit an inspection plan first and read its result before proposing execution.",
    "- Per server turn: at most one inspect, then at most one execute; no plan after execute. Inspect and execute must share the same trusted turnKey, associated by the host, not supplied in model arguments. Inspection is optional when the required workbook facts are already known.",
    "- Plan within these limits for serialized server inspection results: 1 MiB per read, 1.5 MiB for the body, nesting depth 64, and 5,000 cells returned to the server model. Byte limits count compact JSON in UTF-8; depth counts the read object as level 1.",
    `- Bound ask / refuse / finalSummary text to ${schema.properties.message.maxLength.toLocaleString("en-US")} Unicode code points. Ask/refuse messages are forwarded as finalSummary; execute summary has the same ${schema.properties.summary.maxLength.toLocaleString("en-US")}-code-point bound.`,
    "- Wait for and read the tool result before claiming success; nothing executes from the plan alone.",
    "- Keep plans bounded and explicit: destructive and replacement fields are not part of V1. Do not invent private routing, session, policy, limit, request, idempotency or fingerprint fields.",
    "",
    "Non-executing reasons: ask, refuse, rejected, cancelled, not_started, invalid, already_executed.",
    "Do not retry after already_executed: an earlier attempt of this same operation may already have changed the workbook. Do not run it again. Read the recorded execution report in the reply and tell the user what it shows (it may be partial or unknown). Do not treat an uncertain result as proof that nothing changed.",
    "",
    `Schema fingerprint: ${workbookPlan["x-workbook-plan-fingerprint"]}.`,
    "",
  ].join("\n");
}

function checkArtifact(path, expected, missingMessage, staleMessage) {
  let current;
  try {
    current = readFileSync(path, "utf8");
  } catch {
    console.error(`${missingMessage}: ${path}`);
    process.exitCode = 1;
    return;
  }
  if (current.replace(/\r\n/g, "\n") !== expected.replace(/\r\n/g, "\n")) {
    console.error(staleMessage);
    process.exitCode = 1;
  }
}

try {
  const workbookPlan = buildWorkbookPlanSchemaTool();
  const output = `${JSON.stringify([...buildPluginSchemas(compileCanonicalSchemas()), workbookPlan], null, 2)}\n`;
  const skillOutput = buildSkill(workbookPlan);
  if (checkOnly) {
    checkArtifact(outputPath, output, "Generated schema artifact is missing", "Generated Excel plugin schemas are stale. Run npm run generate:plugin-schemas.");
    checkArtifact(skillPath, skillOutput, "Generated workbook plan skill is missing", "Generated workbook plan skill is stale. Run npm run generate:plugin-schemas.");
  } else {
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, output, "utf8");
    mkdirSync(dirname(skillPath), { recursive: true });
    writeFileSync(skillPath, skillOutput, "utf8");
    console.log(`Wrote ${outputPath}`);
    console.log(`Wrote ${skillPath}`);
  }
} finally {
  rmSync(tempDir, { recursive: true, force: true });
}
