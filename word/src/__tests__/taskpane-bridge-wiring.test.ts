import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";

const read = (relativePath: string): string =>
  readFileSync(path.resolve(process.cwd(), relativePath), "utf8");

describe("Word taskpane lifecycle integration contract", () => {
  it("delegates pagehide and visibility behavior to the tested lifecycle controller", () => {
    const taskpane = read("src/taskpane/taskpane.ts");
    expect(taskpane).toContain('from "../officeBridgeLifecycle"');
    expect(taskpane).toContain("new OfficeBridgeLifecycleController");
    expect(taskpane).toContain("bridgeLifecycle.start()");
    expect(taskpane).toContain("bridgeLifecycle.register(");
    expect(taskpane).toContain("Office.addin.onVisibilityModeChanged");
    expect(taskpane).toContain("Office.VisibilityMode.hidden");
  });

  it("shows a dedicated bridge state and copy control", () => {
    const html = read("src/taskpane/taskpane.html");
    const css = read("src/taskpane/omnia-pane.css") + read("src/taskpane/taskpane.css");
    expect(html).toContain('id="bridge-status"');
    expect(html).toContain('id="bridge-session-id"');
    expect(html).toContain('id="copy-bridge-session"');
    expect(css).toContain(".bridge-strip");
    expect(css).toContain(".bridge-session-id");
  });

  it("reaches the Sidecar bridge on 127.0.0.1 with a remembered grant and no API key", () => {
    const webpack = read("webpack.config.js");
    expect(webpack).not.toContain("devServer");
    expect(webpack).not.toContain("office-addin-dev-certs");
    const taskpane = read("src/taskpane/taskpane.ts");
    expect(taskpane).toContain("connector.discover()");
    expect(taskpane).toContain("grant: () => connector.grant");
    expect(taskpane).toContain("requestHeaders: chatHeaders");
    expect(taskpane).not.toMatch(/apiKey|Authorization/);
    const settings = read("src/settings.ts");
    expect(settings).not.toContain("apiKey:");
    const html = read("src/taskpane/taskpane.html");
    expect(html).toContain('id="desk-connect-button"');
    expect(html).not.toContain("api-key-input");
    expect(html).not.toContain("Omnia");
  });

  it("keeps Office unregister commands out of the taskpane production entry point", () => {
    const taskpane = read("src/taskpane/taskpane.ts");
    expect(taskpane).not.toContain("stop:sideload");
    expect(taskpane).not.toContain("office-addin-debugging stop");
  });

  it("keeps cached-result semantics in executable bridge tests", () => {
    const bridgeTests = read("src/__tests__/office-bridge.test.ts");
    expect(bridgeTests).toContain("retries a cached result without re-executing a mutating ticket");
    expect(bridgeTests).toContain("drops a terminally rejected cached result and resumes polling");
  });

  it("keeps lifecycle behavior in executable controller tests", () => {
    const lifecycleTests = read("src/__tests__/office-bridge-lifecycle.test.ts");
    expect(lifecycleTests).toContain("deduplicates concurrent initial starts");
    expect(lifecycleTests).toContain("stops when hidden and restarts when visible");
    expect(lifecycleTests).toContain("pagehide stops the bridge and removes the visibility handler");
    expect(lifecycleTests).toContain("reports start failures");
  });
});
