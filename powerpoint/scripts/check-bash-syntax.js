#!/usr/bin/env node
"use strict";

const { readdirSync } = require("fs");
const { join } = require("path");
const { spawnSync } = require("child_process");
const { resolveBash } = require("./resolve-bash.js");

const scriptsDir = __dirname;
const shellScripts = readdirSync(scriptsDir)
  .filter((name) => name.endsWith(".sh"))
  .sort();
let bash;
try {
  bash = resolveBash();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}

for (const name of shellScripts) {
  // Git Bash understands drive-letter paths with forward slashes. Backslashes
  // are shell escapes, not portable path separators.
  const script = join(scriptsDir, name).replace(/\\/g, "/");
  const result = spawnSync(bash, ["-n", script], { encoding: "utf8", shell: false });
  if (result.error || result.status !== 0) {
    process.stderr.write(result.error?.message || result.stderr || `bash -n failed for ${name}\n`);
    process.exit(result.status ?? 1);
  }
}
console.log(`Validated ${shellScripts.length} bash scripts.`);
