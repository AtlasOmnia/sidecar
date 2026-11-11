#!/usr/bin/env node
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(scriptDir, "..");
const outputPath = join(repoRoot, "plugins", "atlas-sidecar-word", "tool_schemas.json");
const checkOnly = process.argv.includes("--check");
const tempDir = mkdtempSync(join(repoRoot, ".schema-build-"));

function compileCanonicalSchemas() {
  const tsc = join(repoRoot, "node_modules", "typescript", "bin", "tsc");
  execFileSync(process.execPath, [
    tsc,
    "src/wordTools.ts",
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
  const compiled = require(join(tempDir, "wordTools.js"));
  if (!Array.isArray(compiled.toolSchemas)) {
    throw new Error("Canonical src/wordTools.ts did not export toolSchemas");
  }
  return compiled.toolSchemas;
}

function buildPluginSchemas(canonicalSchemas) {
  return canonicalSchemas.map((wrapper) => {
    if (wrapper?.type !== "function" || !wrapper.function) {
      throw new Error("Unexpected canonical OpenAI tool schema shape");
    }
    const schema = structuredClone(wrapper.function);
    schema.name = `sidecar_word_${schema.name}`;
    schema.parameters ??= { type: "object", properties: {} };
    schema.parameters.additionalProperties = false;
    return schema;
  });
}

try {
  const output = `${JSON.stringify(buildPluginSchemas(compileCanonicalSchemas()), null, 2)}\n`;
  if (checkOnly) {
    let current;
    try {
      current = readFileSync(outputPath, "utf8");
    } catch {
      console.error(`Generated schema artifact is missing: ${outputPath}`);
      process.exitCode = 1;
      current = null;
    }
    if (current !== null && current !== output) {
      console.error("Generated Word plugin schemas are stale. Run npm run generate:plugin-schemas.");
      process.exitCode = 1;
    }
  } else {
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, output, "utf8");
    console.log(`Wrote ${outputPath}`);
  }
} finally {
  rmSync(tempDir, { recursive: true, force: true });
}
