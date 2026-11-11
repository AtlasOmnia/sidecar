import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";

const read = (relativePath: string): string =>
  readFileSync(path.resolve(process.cwd(), relativePath), "utf8");

describe("PowerPoint taskpane companion bridge wiring", () => {
  it("shows a dedicated bridge state and copy control", () => {
    const html = read("src/taskpane/taskpane.html");
    const css = read("src/taskpane/omnia-pane.css") + read("src/taskpane/taskpane.css");
    expect(html).toContain('id="bridge-status"');
    expect(html).toContain('id="bridge-session-id"');
    expect(html).toContain('id="copy-bridge-session"');
    expect(css).toContain(".bridge-strip");
    expect(css).toContain(".bridge-session-id");
  });

  it("bridge status never exposes keys or secrets in HTML", () => {
    const html = read("src/taskpane/taskpane.html");
    // Session key, api key, and other secrets must never appear in HTML
    expect(html).not.toMatch(/sessionKey|session-?key|plugin-?key/i);
    // The session ID display is a code element with a specific id — only that
    const codePattern = /<code[^>]*id="bridge-session-id"[^>]*>/g;
    const matches = html.match(codePattern);
    expect(matches).not.toBeNull();
    expect(matches!.length).toBe(1);
  });

  it("starts after Office readiness and delegates pane lifecycle changes", () => {
    const source = read("src/taskpane/taskpane.ts");
    expect(source).toContain('from "../officeBridgeLifecycle"');
    expect(source).toContain('from "../officeBridgePowerPoint"');
    expect(source).toMatch(/Office\.onReady\([\s\S]*?bridgeLifecycle\.start\(\)/);
    expect(source).toContain("bridgeLifecycle.register(");
    expect(source).toContain("Office.addin.onVisibilityModeChanged");
  });

  it("executes through validateToolCall and executePowerPointToolParsed", () => {
    const bridgeSource = read("src/officeBridgePowerPoint.ts");
    // The executeValidatedPowerPointBridgeTool function must call
    // validateToolCall first, then only on success call executePowerPointToolParsed
    const funcStart = bridgeSource.indexOf("export async function executeValidatedPowerPointBridgeTool");
    const nextFunc = bridgeSource.indexOf("export function createPowerPointCompanionBridge", funcStart + 1);
    const funcBody = bridgeSource.slice(funcStart, nextFunc);
    expect(funcBody).toContain("validateToolCall");
    expect(funcBody).toContain("executePowerPointToolParsed");
    // validateToolCall must be called before executePowerPointToolParsed in the function body
    const valIdx = funcBody.indexOf("validateToolCall(name,");
    const execIdx = funcBody.indexOf("executePowerPointToolParsed(name,");
    expect(valIdx).toBeLessThan(execIdx);
  });

  it("preserves capability-gated unsupported responses", () => {
    const toolsSource = read("src/powerPointTools.ts");
    // add_speaker_notes must return capability-gated unsupported
    expect(toolsSource).toContain("add_speaker_notes");
    expect(toolsSource).toMatch(/ok:\s*false/);
    // duplicate_slide must return capability-gated unsupported
    expect(toolsSource).toContain("duplicate_slide");
    // add_image must be capability-gated
    expect(toolsSource).toContain("add_image");
    expect(toolsSource).toContain("available: false");
  });
});
