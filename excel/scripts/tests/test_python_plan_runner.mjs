#!/usr/bin/env node

import assert from "node:assert/strict";
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, unlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { delimiter, dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const runnerPath = join(repositoryRoot, "scripts", "run-python-tests.mjs");

function fakeInterpreterSource() {
  return `#!${process.execPath}
import { appendFileSync, unlinkSync } from "node:fs";

const args = process.argv.slice(2);
if (process.env.FAKE_LOG) {
  appendFileSync(process.env.FAKE_LOG, JSON.stringify(args) + "\\n");
}

if (args[0] === "-c") {
  if (process.env.FAKE_DELETE_AFTER_PROBE === "1") {
    unlinkSync(process.argv[1]);
  }
  process.stdout.write("HERMES_PYTHON_OK\\n");
  process.exit(0);
}

if (args[0] === "-m" && args[1] === "pytest") {
  if (process.env.FAKE_OUTCOME === "SIGTERM") {
    process.kill(process.pid, "SIGTERM");
  }
  process.exit(Number(process.env.FAKE_OUTCOME ?? "0"));
}

process.exit(Number(process.env.FAKE_DIRECT_OUTCOME ?? "0"));
`;
}

function createExecutable(path) {
  writeFileSync(path, fakeInterpreterSource(), { encoding: "utf8", mode: 0o755 });
  chmodSync(path, 0o755);
}

function windowsSpawnDoubleSource() {
  // Windows cannot execute a POSIX shebang fake. Inject only the child-process
  // seam in the runner child; real interpreter integration has its own gates.
  return `import childProcess from "node:child_process";
import { syncBuiltinESMExports } from "node:module";
import { appendFileSync, existsSync, unlinkSync } from "node:fs";
childProcess.spawnSync = (command, args) => {
  appendFileSync(process.env.FAKE_ATTEMPT_LOG, JSON.stringify(command) + "\\n");
  if (!existsSync(command)) return { status: null, error: new Error("ENOENT"), stdout: "" };
  appendFileSync(process.env.FAKE_LOG, JSON.stringify(args) + "\\n");
  if (args[0] === "-c") {
    if (process.env.FAKE_DELETE_AFTER_PROBE === "1") unlinkSync(command);
    return { status: 0, signal: null, stdout: "HERMES_PYTHON_OK\\n" };
  }
  if (args[0] === "-m" && args[1] === "pytest") {
    if (process.env.FAKE_OUTCOME === "SIGTERM") return { status: null, signal: "SIGTERM" };
    return { status: Number(process.env.FAKE_OUTCOME), signal: null };
  }
  return { status: Number(process.env.FAKE_DIRECT_OUTCOME), signal: null };
};
syncBuiltinESMExports();
`;
}

function runRunner({ runnerArgs, outcome = "0", directOutcome = "0", configuredPath = "fake", fallback = false, deleteAfterProbe = false }) {
  const tempRoot = mkdtempSync(join(tmpdir(), "hermes-python-plan-runner-"));
  const logPath = join(tempRoot, "argv.log");
  const attemptLogPath = join(tempRoot, "attempts.log");
  const configuredInterpreter = configuredPath === "fake" ? join(tempRoot, "configured-python") : configuredPath;
  if (configuredPath === "fake") {
    createExecutable(configuredInterpreter);
  }

  let path = process.env.PATH ?? "";
  if (fallback) {
    const fallbackDirectory = join(tempRoot, "bin");
    const fallbackInterpreter = join(fallbackDirectory, "python3");
    mkdirSync(fallbackDirectory, { recursive: true });
    createExecutable(fallbackInterpreter);
    path = `${fallbackDirectory}${delimiter}${path}`;
  }

  const environment = {
    ...process.env,
    FAKE_LOG: logPath,
    FAKE_ATTEMPT_LOG: attemptLogPath,
    FAKE_OUTCOME: outcome,
    FAKE_DIRECT_OUTCOME: directOutcome,
    FAKE_DELETE_AFTER_PROBE: deleteAfterProbe ? "1" : "0",
    PATH: path,
  };
  if (configuredPath === undefined) {
    delete environment.PYTHON;
  } else {
    environment.PYTHON = configuredInterpreter;
  }

  try {
    const preloadArgs = [];
    if (process.platform === "win32") {
      const preload = join(tempRoot, "spawn-double.mjs");
      writeFileSync(preload, windowsSpawnDoubleSource());
      preloadArgs.push("--import", pathToFileURL(preload).href);
    }
    const result = spawnSync(process.execPath, [...preloadArgs, runnerPath, ...runnerArgs], {
      cwd: repositoryRoot,
      env: environment,
      encoding: "utf8",
      timeout: 5000,
    });
    const calls = existsSync(logPath)
      ? readFileSync(logPath, "utf8").trim().split("\n").filter(Boolean).map((line) => JSON.parse(line))
      : [];
    const attempts = existsSync(attemptLogPath)
      ? readFileSync(attemptLogPath, "utf8").trim().split("\n").filter(Boolean).map((line) => JSON.parse(line))
      : [];
    return { result, calls, attempts };
  } finally {
    rmSync(tempRoot, { recursive: true, force: true });
  }
}

function assertPytestCall(calls, targets) {
  const pytestCalls = calls.filter((args) => args[0] === "-m" && args[1] === "pytest");
  assert.equal(pytestCalls.length, 1, `expected one pytest child invocation; saw ${JSON.stringify(calls)}`);
  assert.deepEqual(pytestCalls[0], ["-m", "pytest", "-q", "-p", "no:cacheprovider", ...targets]);
}

function testValidPytestInvocation() {
  const targets = ["scripts/tests/first_plan_test.py", "scripts/tests/second_plan_test.py"];
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", ...targets] });
  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.signal, null);
  assert.equal(calls.length, 2);
  assert.equal(calls[0][0], "-c");
  assert.match(calls[0][1], /sys\.version_info/);
  assertPytestCall(calls, targets);
}

function testMissingTargetStatusAndArgvArePreserved() {
  const target = "scripts/tests/missing_plan_test.py";
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", target], outcome: "4" });
  assert.equal(result.status, 4, result.stderr);
  assertPytestCall(calls, [target]);
}

function testNoTestsStatusAndArgvArePreserved() {
  const target = "scripts/tests/no_tests_plan_test.py";
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", target], outcome: "5" });
  assert.equal(result.status, 5, result.stderr);
  assertPytestCall(calls, [target]);
}

function testCollectionFailureStatusAndArgvArePreserved() {
  const target = "scripts/tests/collection_failure_plan_test.py";
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", target], outcome: "2" });
  assert.equal(result.status, 2, result.stderr);
  assertPytestCall(calls, [target]);
}

function testChildFailureStatusAndArgvArePreserved() {
  const target = "scripts/tests/failing_plan_test.py";
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", target], outcome: "7" });
  assert.equal(result.status, 7, result.stderr);
  assertPytestCall(calls, [target]);
}

function testSignalFailureIsNonzeroAndUsesPytestInvocation() {
  const target = "scripts/tests/signaled_plan_test.py";
  const { result, calls } = runRunner({ runnerArgs: ["--pytest", target], outcome: "SIGTERM" });
  assert.equal(result.status, 1, result.stderr);
  assertPytestCall(calls, [target]);
}

function testConfiguredPythonProbeFailureDoesNotFallbackInPytestMode() {
  const tempRoot = mkdtempSync(join(tmpdir(), "hermes-python-plan-runner-missing-"));
  const missingInterpreter = join(tempRoot, "does-not-exist");
  try {
    const { result, calls, attempts } = runRunner({
      runnerArgs: ["--pytest", "scripts/tests/should-not-run.py"],
      configuredPath: missingInterpreter,
      fallback: true,
    });
    assert.equal(result.status, 1, result.stderr);
    assert.deepEqual(calls, [], "pytest mode must not fall back after configured PYTHON failure");
    if (process.platform === "win32") {
      assert.deepEqual(attempts, [missingInterpreter], "failed fallback attempts must not disappear");
    }
  } finally {
    rmSync(tempRoot, { recursive: true, force: true });
  }
}

function testChildSpawnErrorIsNonzeroAfterSuccessfulProbe() {
  const target = "scripts/tests/spawn_error_plan_test.py";
  const { result, calls } = runRunner({
    runnerArgs: ["--pytest", target],
    deleteAfterProbe: true,
  });
  assert.equal(result.status, 1, result.stderr);
  assert.match(result.stderr, /Unable to run/);
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], "-c");
  assert.match(calls[0][1], /sys\.version_info/);
}

function testDirectScriptModeRetainsOriginalArgvAndMultipleTargetSemantics() {
  const targets = ["scripts/test_excel_no_key_proxy.py", "scripts/test_proxy_passthrough.py"];
  const { result, calls } = runRunner({ runnerArgs: targets });
  assert.equal(result.status, 0, result.stderr);
  assert.equal(calls.length, 3);
  assert.equal(calls[0][0], "-c");
  assert.deepEqual(calls.slice(1), targets.map((target) => [target]));
}

const tests = [
  ["real configured interpreter integration", () => {
    const tempRoot = mkdtempSync(join(tmpdir(), "hermes-python-plan-runner-real-"));
    try {
      const target = join(tempRoot, "test_real_runner.py");
      writeFileSync(target, "def test_real_runner():\n    assert 1 == 1\n");
      const result = spawnSync(process.execPath, [runnerPath, "--pytest", target], {
        cwd: repositoryRoot,
        env: { ...process.env, PYTHON: process.env.PYTHON ?? (process.platform === "win32" ? "python" : "python3") },
        encoding: "utf8",
        timeout: 10000,
      });
      assert.equal(result.status, 0, result.stdout + result.stderr);
      assert.match(result.stdout, /1 passed/);
    } finally {
      rmSync(tempRoot, { recursive: true, force: true });
    }
  }],
  ["valid pytest invocation", testValidPytestInvocation],
  ["missing target propagation", testMissingTargetStatusAndArgvArePreserved],
  ["no-tests propagation", testNoTestsStatusAndArgvArePreserved],
  ["collection failure propagation", testCollectionFailureStatusAndArgvArePreserved],
  ["child failure propagation", testChildFailureStatusAndArgvArePreserved],
  ["signal failure propagation", testSignalFailureIsNonzeroAndUsesPytestInvocation],
  ["configured interpreter fail-closed", testConfiguredPythonProbeFailureDoesNotFallbackInPytestMode],
  ["child spawn error", testChildSpawnErrorIsNonzeroAfterSuccessfulProbe],
  ["direct script compatibility", testDirectScriptModeRetainsOriginalArgvAndMultipleTargetSemantics],
];

let failures = 0;
for (const [name, test] of tests) {
  try {
    test();
    console.log(`PASS ${name}`);
  } catch (error) {
    failures += 1;
    console.error(`FAIL ${name}`);
    console.error(error instanceof Error ? error.stack : error);
  }
}

if (failures > 0) {
  console.error(`${failures} of ${tests.length} runner regression tests failed.`);
  process.exitCode = 1;
} else {
  console.log(`${tests.length} runner regression tests passed.`);
}
