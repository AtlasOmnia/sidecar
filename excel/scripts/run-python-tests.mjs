#!/usr/bin/env node
import { spawnSync } from "node:child_process";

const requestedTests = process.argv.slice(2);
const pytestMode = requestedTests[0] === "--pytest";
const tests = pytestMode ? requestedTests.slice(1) : requestedTests;
if (tests.length === 0) {
  console.error("Usage: node scripts/run-python-tests.mjs [--pytest] <test.py> [...]");
  process.exit(2);
}

const fallbackCandidates = process.platform === "win32"
  ? [["python", []], ["py", ["-3"]], ["python3", []]]
  : [["python3", []], ["python", []]];
const configuredPython = process.env.PYTHON?.trim();
const candidates = pytestMode
  ? (configuredPython ? [[configuredPython, []]] : fallbackCandidates)
  : (configuredPython
    ? [[configuredPython, []], ...fallbackCandidates]
    : fallbackCandidates);

let python = null;
for (const [command, prefixArgs] of candidates) {
  const probe = spawnSync(
    command,
    [...prefixArgs, "-c", "import sys; assert sys.version_info.major == 3; print('HERMES_PYTHON_OK')"],
    {
      encoding: "utf8",
      shell: false,
    },
  );
  if (probe.status === 0 && probe.stdout?.includes("HERMES_PYTHON_OK")) {
    python = { command, prefixArgs };
    break;
  }
}

if (!python) {
  console.error("Python 3 is required but no supported interpreter was found.");
  process.exit(1);
}

if (pytestMode) {
  const result = spawnSync(
    python.command,
    [...python.prefixArgs, "-m", "pytest", "-q", "-p", "no:cacheprovider", ...tests],
    {
      cwd: process.cwd(),
      stdio: "inherit",
      shell: false,
    },
  );
  if (result.error) {
    console.error(`Unable to run pytest: ${result.error.message}`);
    process.exit(1);
  }
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
} else {
  for (const test of tests) {
    const result = spawnSync(python.command, [...python.prefixArgs, test], {
      cwd: process.cwd(),
      stdio: "inherit",
      shell: false,
    });
    if (result.error) {
      console.error(`Unable to run ${test}: ${result.error.message}`);
      process.exit(1);
    }
    if (result.status !== 0) {
      process.exit(result.status ?? 1);
    }
  }
}
