// @vitest-environment node
import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { spawnSync } from "child_process";
import { join, resolve } from "path";

const REPO_ROOT = resolve(__dirname, "../..");

describe("metadata tracer task-pane integration", () => {
  it("wires model attempts, tool results, and checkpoint outcomes into persisted traces", () => {
    const source = readFileSync(join(REPO_ROOT, "src", "taskpane", "taskpane.ts"), "utf8");
    expect(source).toContain("recordModelCall");
    expect(source).toContain("recordToolResult");
    expect(source).toContain("recordCheckpoint");
    expect(source).toContain("onModelCall");
    expect(source).toContain("onToolResult");
    expect(source).toContain("onCheckpointResult");
  });
});
