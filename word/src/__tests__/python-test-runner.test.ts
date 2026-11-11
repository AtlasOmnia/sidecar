import { describe, expect, it } from "vitest";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { pathToFileURL } from "node:url";

const runnerUrl = pathToFileURL(path.resolve(process.cwd(), "scripts/run-python-tests.mjs")).href;

function probe(source: string): void {
  const result = spawnSync(process.execPath, ["--input-type=module", "--eval", `
import assert from "node:assert/strict";
import { pythonCandidates, runPythonTests } from ${JSON.stringify(runnerUrl)};
const ok = { status: 0, stdout: "HERMES_PYTHON_OK\\n" };
${source}
`], { encoding: "utf8", shell: false });
  expect(result.status, result.stderr || result.stdout).toBe(0);
}

describe("cross-platform Python test runner", () => {
  it("preserves macOS python3 preference and Windows fallbacks", () => {
    probe(`
assert.deepEqual(pythonCandidates("darwin", {}), [["python3", []], ["python", []]]);
assert.deepEqual(pythonCandidates("win32", {}), [["python", []], ["py", ["-3"]], ["python3", []]]);
assert.deepEqual(pythonCandidates("win32", { PYTHON: " C:/Python With Spaces/python.exe " }), [["C:/Python With Spaces/python.exe", []]]);
`);
  });

  it("rejects Store stubs and uses the launcher prefix for pytest", () => {
    probe(`
const calls = [];
const code = runPythonTests(["--pytest", "test.py", "-k", "safe"], {
  platform: "win32", env: {}, cwd: "authorized-repo",
  spawnSyncFn: (command, args, options) => {
    calls.push({ command, args, options });
    if (args[0] === "-c") return { status: 9009, stdout: "" };
    if (args.includes("-c")) return ok;
    return { status: 0 };
  },
});
assert.equal(code, 0);
assert.deepEqual(calls.map(call => call.command), ["python", "py", "py"]);
assert.deepEqual(calls[2].args, ["-3", "-m", "pytest", "-q", "-p", "no:cacheprovider", "test.py", "-k", "safe"]);
assert.equal(calls[2].options.cwd, "authorized-repo");
assert.equal(calls[2].options.shell, false);
assert.equal(calls[2].options.stdio, "inherit");
`);
  });

  it("fails closed for an invalid explicit override without fallback", () => {
    probe(`
const calls = [];
assert.equal(runPythonTests(["--pytest", "test.py"], {
  env: { PYTHON: "broken" }, platform: "win32", logError: () => {},
  spawnSyncFn: (command) => { calls.push(command); return { status: 0, stdout: "not Python" }; },
}), 1);
assert.deepEqual(calls, ["broken"]);
`);
  });

  it("propagates pytest failures without trying another interpreter", () => {
    probe(`
let count = 0;
assert.equal(runPythonTests(["--pytest", "test.py"], {
  env: { PYTHON: "chosen" }, logError: () => {},
  spawnSyncFn: () => (++count === 1 ? ok : { status: 17 }),
}), 17);
assert.equal(count, 2);
`);
  });

  it("does not report spawn errors or signals as a pass", () => {
    probe(`
for (const failure of [{ status: null, error: new Error("spawn failed") }, { status: null, signal: "SIGTERM" }]) {
  let count = 0;
  assert.equal(runPythonTests(["--pytest", "test.py"], {
    env: { PYTHON: "chosen" }, logError: () => {},
    spawnSyncFn: () => (++count === 1 ? ok : failure),
  }), 1);
}
`);
  });

  it("requires tests before spawning anything", () => {
    probe(`
for (const args of [[], ["--pytest"]]) {
  assert.equal(runPythonTests(args, {
    logError: () => {}, spawnSyncFn: () => { throw new Error("unexpected spawn"); },
  }), 2);
}
`);
  });

  it("runs standalone scripts sequentially and stops at the first failure", () => {
    probe(`
const calls = [];
assert.equal(runPythonTests(["one.py", "two.py", "three.py"], {
  env: { PYTHON: "chosen" },
  spawnSyncFn: (_command, args) => {
    calls.push(args);
    return args.includes("-c") ? ok : { status: args[0] === "two.py" ? 9 : 0 };
  },
}), 9);
assert.deepEqual(calls.slice(1), [["one.py"], ["two.py"]]);
`);
  });
});
