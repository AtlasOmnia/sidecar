import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { describe, expect, it } from "vitest";

function findPythonExecutable(): string {
  const locator = process.platform === "win32" ? "where.exe" : "which";
  const candidates = process.env.PYTHON?.trim()
    ? [process.env.PYTHON.trim()]
    : process.platform === "win32"
      ? ["python", "py", "python3"]
      : ["python3", "python"];

  for (const candidate of candidates) {
    if (candidate.includes("/") || candidate.includes("\\")) {
      return candidate;
    }
    const located = spawnSync(locator, [candidate], { encoding: "utf8", shell: false });
    if (located.status === 0 && located.stdout) {
      const executable = located.stdout.trim().split(/\r?\n/)[0];
      if (executable) {
        return executable;
      }
    }
  }

  throw new Error("A Python executable is required for this regression test.");
}

describe("run-python-tests runner", () => {
  it("uses a nonempty PYTHON executable before PATH fallbacks", () => {
    const directory = mkdtempSync(join(tmpdir(), "hermes-python override-"));
    const testPath = join(directory, "sentinel test.py");
    const runnerPath = resolve(process.cwd(), "scripts/run-python-tests.mjs");
    const explicitPython = findPythonExecutable();

    writeFileSync(testPath, "print('PYTHON_OVERRIDE_SENTINEL')\n");

    try {
      const result = spawnSync(process.execPath, [runnerPath, testPath], {
        encoding: "utf8",
        env: {
          ...process.env,
          PATH: directory,
          PYTHON: explicitPython,
        },
        shell: false,
      });

      expect(result.status).toBe(0);
      expect(result.stdout).toContain("PYTHON_OVERRIDE_SENTINEL");
    } finally {
      rmSync(directory, { force: true, recursive: true });
    }
  }, 120_000);
});
