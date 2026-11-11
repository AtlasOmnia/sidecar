import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";

function readSource(relativePath: string): string {
  return readFileSync(path.resolve(process.cwd(), relativePath), "utf8");
}

describe("PowerPoint taskpane lifecycle integration contract", () => {
  it("delegates pagehide and visibility behavior to the tested lifecycle controller", () => {
    const taskpane = readSource("src/taskpane/taskpane.ts");
    expect(taskpane).toContain('from "../officeBridgeLifecycle"');
    expect(taskpane).toContain("new OfficeBridgeLifecycleController");
    expect(taskpane).toContain("bridgeLifecycle.start()");
    expect(taskpane).toContain("bridgeLifecycle.register(");
    expect(taskpane).toContain("Office.addin.onVisibilityModeChanged");
    expect(taskpane).toContain("Office.VisibilityMode.hidden");
  });

  it("keeps Office unregister commands out of the taskpane production entry point", () => {
    const taskpane = readSource("src/taskpane/taskpane.ts");
    expect(taskpane).not.toContain("stop:sideload");
    expect(taskpane).not.toContain("office-addin-debugging stop");
  });

  it("keeps cached-result semantics in executable bridge tests", () => {
    const bridgeTests = readSource("src/__tests__/office-bridge.test.ts");
    expect(bridgeTests).toContain("retries a transiently failed cached result without re-executing the mutation");
    expect(bridgeTests).toContain("drops a terminally rejected cached result and resumes polling without re-execution");
  });

  it("keeps lifecycle behavior in executable controller tests", () => {
    const lifecycleTests = readSource("src/__tests__/office-bridge-lifecycle.test.ts");
    expect(lifecycleTests).toContain("deduplicates concurrent initial starts");
    expect(lifecycleTests).toContain("stops when hidden and restarts when visible");
    expect(lifecycleTests).toContain("pagehide stops the bridge and removes the visibility handler");
    expect(lifecycleTests).toContain("reports start failures");
  });
});
