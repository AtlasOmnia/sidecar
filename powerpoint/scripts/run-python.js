#!/usr/bin/env node
"use strict";

const { spawnSync } = require("child_process");

function resolvePython({ platform = process.platform, env = process.env, spawn = spawnSync } = {}) {
  // PYTHON is one executable path, not a shell command. Never silently replace
  // a configured venv: doing so could execute tests with the wrong dependencies.
  const configured = env.PYTHON?.trim();
  const candidates = configured
    ? [[configured, []]]
    : platform === "win32"
      ? [["python", []], ["py", ["-3"]], ["python3", []]]
      : [["python3", []], ["python", []]];
  for (const [command, prefixArgs] of candidates) {
    const probe = spawn(command, [...prefixArgs, "-c",
      "import sys; assert sys.version_info.major == 3; print('HERMES_PYTHON_OK')"], {
      encoding: "utf8", shell: false, timeout: 5000, env,
    });
    if (!probe.error && probe.status === 0 && probe.stdout?.trim() === "HERMES_PYTHON_OK") {
      return { command, prefixArgs };
    }
  }
  throw new Error("Python 3 is required; set PYTHON to a working interpreter or project-local venv executable.");
}

function runPython(args, options = {}) {
  try {
    const { command, prefixArgs } = resolvePython(options);
    const spawn = options.spawn ?? spawnSync;
    const result = spawn(command, [...prefixArgs, ...args], {
      cwd: process.cwd(), stdio: "inherit", shell: false, env: options.env ?? process.env,
    });
    if (result.error) console.error(`Unable to run Python: ${result.error.message}`);
    return result.error ? 1 : result.status ?? 1;
  } catch (error) {
    console.error(error.message);
    return 1;
  }
}

module.exports = { resolvePython, runPython };
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error("Usage: node scripts/run-python.js <Python arguments...>");
    process.exit(2);
  }
  process.exit(runPython(args));
}
