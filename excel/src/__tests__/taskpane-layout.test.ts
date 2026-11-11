import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const css = ["omnia-pane.css", "taskpane.css"].map(name => readFileSync(resolve(__dirname, "../taskpane", name), "utf8")).join("\n");
const html = readFileSync(resolve(__dirname, "../taskpane/taskpane.html"), "utf8");

function rule(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const matches = [...css.matchAll(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, "gs"))];
  // Media queries override spacing, not the shared flex/overflow contract.
  return matches.map(match => match[1]).join("\n");
}

describe("responsive Office task-pane composer", () => {
  it("does not impose a width wider than a narrow Office pane", () => {
    expect(rule("body")).not.toMatch(/min-width:\s*320px/);
    expect(rule(".app-shell")).toMatch(/width:\s*100%/);
    expect(rule(".app-shell")).toMatch(/min-width:\s*0/);
    expect(css).toMatch(/overflow-x:\s*hidden/);
    expect(css).toMatch(/@media\s*\(max-width:\s*360px\)/);
  });

  it("keeps the conversation scrollable while the full composer remains visible", () => {
    expect(rule(".app-shell")).toMatch(/display:\s*flex/);
    expect(rule(".app-shell")).toMatch(/flex-direction:\s*column/);
    expect(rule(".messages")).toMatch(/flex:\s*1\s+1\s+auto/);
    expect(rule(".messages")).toMatch(/min-height:\s*0/);
    expect(rule(".composer")).toMatch(/flex:\s*0\s+0\s+auto/);
    expect(rule(".composer")).toMatch(/width:\s*100%/);
  });

  it("provides a practical multiline prompt area without clipping controls", () => {
    expect(html).toMatch(/id="prompt-input"[^>]*rows="3"/);
    expect(rule("textarea")).toMatch(/min-height:\s*82px/);
    expect(rule("textarea")).toMatch(/max-height:\s*min\(30vh,\s*180px\)/);
    expect(rule("textarea")).toMatch(/overflow-y:\s*auto/);
  });
});
