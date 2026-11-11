#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";

export function pythonCandidates(platform, env) {
  const configured = env.PYTHON?.trim();
  if (configured) return [[configured, []]];
  return platform === "win32"
    ? [["python", []], ["py", ["-3"]], ["python3", []]]
    : [["python3", []], ["python", []]];
}

export function runPythonTests(requestedTests, {
  platform = process.platform,
  env = process.env,
  cwd = process.cwd(),
  spawnSyncFn = spawnSync,
  logError = console.error,
} = {}) {
  const pytestMode = requestedTests[0] === "--pytest";
  const tests = pytestMode ? requestedTests.slice(1) : requestedTests;
  if (tests.length === 0) {
    logError("Usage: node scripts/run-python-tests.mjs [--pytest] <test.py> [...]");
    return 2;
  }

  let python = null;
  for (const [command, prefixArgs] of pythonCandidates(platform, env)) {
    const probe = spawnSyncFn(command, [
      ...prefixArgs, "-c",
      "import sys; assert sys.version_info.major == 3; print('HERMES_PYTHON_OK')",
    ], { cwd, env, encoding: "utf8", shell: false });
    if (!probe.error && probe.status === 0 && probe.stdout?.includes("HERMES_PYTHON_OK")) {
      python = { command, prefixArgs };
      break;
    }
  }
  if (!python) {
    logError("Python 3 is required but no supported interpreter was found. Set PYTHON to a Python 3 executable.");
    return 1;
  }

  const invocations = pytestMode
    ? [["-m", "pytest", "-q", "-p", "no:cacheprovider", ...tests]]
    : tests.map((test) => [test]);
  for (const args of invocations) {
    const result = spawnSyncFn(python.command, [...python.prefixArgs, ...args], {
      cwd, env, stdio: "inherit", shell: false,
    });
    if (result.error) {
      logError(`Unable to run Python tests: ${result.error.message}`);
      return 1;
    }
    if (result.status !== 0) return result.status ?? 1;
  }
  return 0;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.exit(runPythonTests(process.argv.slice(2)));
}
